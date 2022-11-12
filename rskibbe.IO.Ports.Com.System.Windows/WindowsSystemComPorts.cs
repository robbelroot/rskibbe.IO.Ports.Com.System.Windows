﻿using rskibbe.IO.Ports.Com.ValueObjects;
using System.Management;
using System.Text.RegularExpressions;

namespace rskibbe.IO.Ports.Com.System.Windows;

public class WindowsSystemComPorts : SystemComPorts, IDisposable
{

    const string QUERY = "SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2 or EventType = 3";

    protected ManagementEventWatcher? watcher;

    public override List<string> ExistingPorts { get; }

    private bool _disposed;

    protected SynchronizationContext? _synchronizationContext;

    public ComWatcherState State { get; protected set; }

    public bool CanStartWatching => State == ComWatcherState.NONE || State == ComWatcherState.STOPPED;

    public bool CanStopWatching => State == ComWatcherState.STOPPED;

    public bool IgnoreComOne { get; set; }

    protected WindowsSystemComPorts()
    {
        ExistingPorts = new List<string>();
        State = ComWatcherState.NONE;
        IgnoreComOne = true;
    }

    public static async Task<WindowsSystemComPorts> BuildAsync(SynchronizationContext? synchronizationContext = null)
    {
        var winPorts = new WindowsSystemComPorts();
        winPorts._synchronizationContext = synchronizationContext ?? SynchronizationContext.Current;
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        winPorts.watcher = new ManagementEventWatcher(QUERY);
        winPorts.watcher.Stopped += winPorts.Watcher_Stopped;
        winPorts.watcher.EventArrived += winPorts.Watcher_EventArrived;
        var usedPortNames = await winPorts.ListUsedPortNamesAsync();
        winPorts.ExistingPorts.AddRange(usedPortNames);
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
        return winPorts;
    }

    /// <exception cref="InvalidOperationException">On invalid state</exception>
    public void StartWatchingPorts()
    {
        if (!CanStartWatching)
            throw new InvalidOperationException($"The watcher cannot be started due to its current state: {State}");
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        watcher!.Start();
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
        State = ComWatcherState.STARTED;
        OnStartedWatchingPorts(EventArgs.Empty);
    }

    protected virtual void OnStartedWatchingPorts(EventArgs e)
       => StartedWatchingPorts?.Invoke(this, e);

    /// <exception cref="InvalidOperationException">On invalid state</exception>
    public void StopWatchingPorts()
    {
        ThrowIfCantStopWatching();
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        watcher!.Stop();
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
        State = ComWatcherState.STOP_REQUESTED;
    }

    /// <exception cref="InvalidOperationException">On invalid state</exception>
    protected void ThrowIfCantStopWatching()
    {
        if (!CanStopWatching)
            throw new InvalidOperationException($"The watcher cannot be stopped due to its current state: {State}");
    }

    protected virtual void OnStoppedWatchingPorts(EventArgs e)
        => StoppedWatchingPorts?.Invoke(this, e);

    /// <exception cref="InvalidOperationException">On invalid state</exception>
    public Task StopWatchingPortsAsync()
    {
        ThrowIfCantStopWatching();
        var tcs = new TaskCompletionSource();
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        StoppedEventHandler? watcherStopped = null;
        watcherStopped = (s, e) =>
        {
            watcher!.Stopped -= watcherStopped;
            tcs.SetResult();
        };
        watcher!.Stopped += watcherStopped;
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
        StopWatchingPorts();
        return tcs.Task;
    }

    private void Watcher_Stopped(object sender, StoppedEventArgs e)
    {
        State = ComWatcherState.STOPPED;
        OnStoppedWatchingPorts(EventArgs.Empty);
    }

    private async void Watcher_EventArrived(object sender, EventArrivedEventArgs e)
    {
        var availablePorts = await ListUsedPortNamesAsync();
        var removedPorts = ExistingPorts.Where(x => !availablePorts.Contains(x)).ToList();
        var addedPorts = availablePorts.Where(x => !ExistingPorts.Contains(x)).ToList();
        foreach (var addedPort in addedPorts)
        {
            ExistingPorts.Add(addedPort);
            OnSystemComPortAdded(new ComPortEventArgs(addedPort));
        }
        foreach (var removedPort in removedPorts)
        {
            ExistingPorts.Remove(removedPort);
            OnSystemComPortRemoved(new ComPortEventArgs(removedPort));
        }
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

    public override async Task<IEnumerable<byte>> ListUsedPortIdsAsync()
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

    public override async Task<IEnumerable<string>> ListUsedPortNamesAsync()
    {
        // change to registry approach / make it like strategy pattern
        // slow AF without scope
        // var scope = "root\\CIMV2";
        var scope = @"\\localhost\root\CIMV2";
        // "SELECT Name,DeviceID FROM Win32_PnPEntity WHERE ClassGuid=`"{4d36e978-e325-11ce-bfc1-08002be10318}`""
        // var query = "SELECT * FROM WIN32_SerialPort";
        var query = "SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0";
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        using var searcher = new ManagementObjectSearcher(scope, query);
        var searchResults = await Task.Run(() => searcher.Get());
        var portObjects = searchResults.Cast<ManagementBaseObject>().ToList();
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
        // DeviceID - COM1
        // PNPDeviceID - "ACPI\\" = non virtual/original win port?
        // Caption - Kommunikationsanschluss (COM1)
        var comRegex = new Regex(@"\bCOM[1-9][0-9]*\b");
        var usedPortNames = new List<string>();
        foreach (var portObject in portObjects)
        {
            if (portObject == null)
                continue;
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
            var captionProp = portObject["Caption"];
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
            if (captionProp == null)
                continue;
            var captionPropValue = captionProp!.ToString();
            if (captionPropValue == null)
                continue;
            var isComPort = captionPropValue!.Contains("COM");
            var noEmulator = !captionPropValue!.Contains("emulator");
            if (isComPort && noEmulator)
            {
                var match = comRegex.Match(captionPropValue);
                if (match.Success)
                {
                    var portName = match.Groups[0].Value;
                    if (portName != "COM1" || portName == "COM1" && !IgnoreComOne)
                        usedPortNames.Add(portName);
                }
            }
        }
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
        portObjects.ForEach(x => x.Dispose());
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
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
#pragma warning disable CA1416 // Plattformkompatibilität überprüfen
                watcher!.Stop();
#pragma warning restore CA1416 // Plattformkompatibilität überprüfen
                watcher.Dispose();
            }

            _disposed = true;
        }
    }

    public event EventHandler? StartedWatchingPorts;

    public event EventHandler? StoppedWatchingPorts;

    public override event EventHandler<ComPortEventArgs>? SystemComPortAdded;

    public override event EventHandler<ComPortEventArgs>? SystemComPortRemoved;
}
