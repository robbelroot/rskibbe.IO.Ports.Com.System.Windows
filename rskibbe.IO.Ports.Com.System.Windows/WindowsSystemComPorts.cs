using rskibbe.IO.Ports.Com.ValueObjects;
using System.Diagnostics;
using System.Management;
using System.Text.RegularExpressions;

using ComPorts = rskibbe.IO.Ports.Com.System.Windows.WindowsSystemComPorts;

namespace rskibbe.IO.Ports.Com.System.Windows;

public class WindowsSystemComPorts : ISystemComPorts, IDisposable
{

    const string QUERY = "SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_SerialPort'";

    protected ManagementEventWatcher? watcher;

    protected List<string> ExistingPorts { get; set; }

    private bool _disposed;

    protected SynchronizationContext? _synchronizationContext;

    protected WindowsSystemComPorts()
    {
        ExistingPorts = new List<string>();
    }

    public static async Task<WindowsSystemComPorts> BuildAsync(SynchronizationContext? synchronizationContext = null)
    {
        var winPorts = new WindowsSystemComPorts();
        winPorts._synchronizationContext = synchronizationContext ?? SynchronizationContext.Current;
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        winPorts.watcher = new ManagementEventWatcher(QUERY);
        // watcher.Stopped += Watcher_Stopped;
        winPorts.watcher.EventArrived += winPorts.Watcher_EventArrived;
        var usedPortNames = await winPorts.ListUsedPortNamesAsync();
        winPorts.ExistingPorts.AddRange(usedPortNames);
        winPorts.watcher.Start();
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
        return winPorts;
    }

    private async void Watcher_EventArrived(object sender, EventArrivedEventArgs e)
    {
        var oldComList = new List<string>(ExistingPorts);
        await RefreshExistingPortsAsync();
        var deviceInserted = oldComList.Count < ExistingPorts.Count;
        var deviceRemoved = oldComList.Count > ExistingPorts.Count;
        if (deviceInserted)
        {
            var insertedPortName = (ExistingPorts).Except(oldComList).First();
            var comPortEventArgs = new ComPortEventArgs(insertedPortName);
            OnSystemComPortAdded(comPortEventArgs);
        }
        else if (deviceRemoved)
        {
            var removedPortName = oldComList.Except(ExistingPorts).First();
            var comPortEventArgs = new ComPortEventArgs(removedPortName);
            OnSystemComPortRemoved(comPortEventArgs);
        }
        else
            Debug.WriteLine($"Not changed - wrong WMI query?");
    }

    protected async Task RefreshExistingPortsAsync()
    {
        ExistingPorts.Clear();
        var usedPortNames = await ListUsedPortNamesAsync();
        ExistingPorts.AddRange(usedPortNames);
    }

    protected virtual void OnSystemComPortAdded(ComPortEventArgs e)
    {
        _synchronizationContext!.Post(new SendOrPostCallback(x =>
        {
            SystemComPortAdded?.Invoke(this, e);
        }), null);
    }

    protected virtual void OnSystemComPortRemoved(ComPortEventArgs e)
    {
        _synchronizationContext!.Post(new SendOrPostCallback(x =>
        {
            SystemComPortRemoved?.Invoke(this, e);
        }), null);
    }

    public async Task<IEnumerable<byte>> ListUsedPortIdsAsync()
    {
        var usedPortNames = await ListUsedPortNamesAsync();
        var usedPortIds = new List<byte>();
        foreach (var portName in usedPortNames)
        {
            portName.ExtractByte(out var id);
            if (id != 0)
                usedPortIds.Add(id);
        }
        return usedPortIds;
    }

    public async Task<IEnumerable<string>> ListUsedPortNamesAsync()
    {
        // change to registry approach / make it like strategy pattern
        // slow AF without scope
        // var scope = "root\\CIMV2";
        var scope = "\\\\localhost\\root\\CIMV2";
        // "SELECT Name,DeviceID FROM Win32_PnPEntity WHERE ClassGuid=`"{4d36e978-e325-11ce-bfc1-08002be10318}`""
        // var query = "SELECT * FROM WIN32_SerialPort";
        var query = "SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0";
        using var searcher = new ManagementObjectSearcher(scope, query);
        var searchResults = await Task.Run(() => searcher.Get());
        var portObjects = searchResults.Cast<ManagementBaseObject>().ToList();
        // DeviceID - COM1
        // PNPDeviceID - "ACPI\\" = non virtual/original win port?
        // Caption - Kommunikationsanschluss (COM1)
        var comRegex = new Regex(@"\bCOM[1-9][0-9]*\b");
        var usedPortNames = new List<string>();
        foreach (var portObject in portObjects)
        {
            var captionProp = portObject["Caption"];
            if (captionProp == null)
                continue;
            var isComPort = captionProp.ToString().Contains("COM");
            var noEmulator = !captionProp.ToString().Contains("emulator");
            if (isComPort && noEmulator)
            {
                var match = comRegex.Match(captionProp.ToString());
                if (match.Success)
                {
                    var portName = match.Groups[0].Value;
                    usedPortNames.Add(portName);
                }
            }
        }
        portObjects.ForEach(x => x.Dispose());
        GC.Collect();
        return usedPortNames;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                watcher.Stop();
                watcher.Dispose();
            }

            _disposed = true;
        }
    }

    public event EventHandler<ComPortEventArgs>? SystemComPortAdded;

    public event EventHandler<ComPortEventArgs>? SystemComPortRemoved;
}
