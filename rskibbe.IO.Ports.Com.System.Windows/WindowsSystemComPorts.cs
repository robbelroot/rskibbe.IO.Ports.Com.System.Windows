using rskibbe.IO.Ports.Com.ValueObjects;
using System.Diagnostics;
using System.Management;
using System.Text.RegularExpressions;

namespace rskibbe.IO.Ports.Com.System.Windows;

public class WindowsSystemComPorts : ISystemComPorts, IDisposable
{

    const string QUERY = "SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_SerialPort'";

    protected ManagementEventWatcher? watcher;

    protected List<string> ExistingPorts { get; set; }

    private bool _disposed;

    public WindowsSystemComPorts()
    {
        ExistingPorts = new List<string>();
        Initialize();
    }

    protected async void Initialize()
    {
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        watcher = new ManagementEventWatcher(QUERY);
        // watcher.Stopped += Watcher_Stopped;
        watcher.EventArrived += Watcher_EventArrived;
        var usedPortNames = await ListUsedPortNamesAsync();
        ExistingPorts.AddRange(usedPortNames);
        watcher.Start();
        OnInitialized(EventArgs.Empty);
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
    }

    protected void OnInitialized(EventArgs e)
    {
        Debug.WriteLine("initialized");
        Initialized?.Invoke(this, e);
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
        SystemComPortAdded?.Invoke(this, e);
    }

    protected virtual void OnSystemComPortRemoved(ComPortEventArgs e)
    {
        SystemComPortRemoved?.Invoke(this, e);
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
        var scope = "root\\CIMV2";
        // "SELECT Name,DeviceID FROM Win32_PnPEntity WHERE ClassGuid=`"{4d36e978-e325-11ce-bfc1-08002be10318}`""
        var query = "SELECT * FROM WIN32_SerialPort";
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
            // var deviceId = "";
            foreach (var prop in portObject.Properties)
            {
                var isCaption = prop.Name == "Caption";
                if (!isCaption)
                    continue;
                var isComPort = prop.Value.ToString().Contains("COM");
                var noEmulator = !prop.Value.ToString().Contains("emulator");
                if (isComPort && noEmulator)
                {
                    var match = comRegex.Match(prop.Value.ToString());
                    if(match.Success)
                    {
                        // could additionally check if port like 300
                        // but okay for now
                        usedPortNames.Add(match.Groups[0].Value);
                        break;
                    }
                }
                //// device id comes before pnpdevice id - sorted a-z
                //// cuz of this, this loop works like this
                //if (prop.Name == "DeviceID")
                //{
                //    deviceId = prop.Value.ToString();
                //}
                //else if (prop.Name == "PNPDeviceID")
                //{
                //    var isWindowsAssignedPort = prop.Value.ToString().StartsWith("ACPI\\");
                //    if (isWindowsAssignedPort)
                //    {
                //        usedPortNames.Add(deviceId);
                //        break;
                //    }
                //}
            }
            break;
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

    public event EventHandler? Initialized;

    public event EventHandler<ComPortEventArgs>? SystemComPortAdded;

    public event EventHandler<ComPortEventArgs>? SystemComPortRemoved;
}
