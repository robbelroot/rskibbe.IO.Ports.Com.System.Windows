using rskibbe.IO.Ports.Com.System.Windows;
using rskibbe.IO.Ports.Com.ValueObjects;
using System.Diagnostics;

namespace TestGui;

public partial class Form1 : Form
{

    WindowsSystemComPorts? _windowsSystemComPorts;

    public Form1()
    {
        InitializeComponent();
        Load += Form1_Load;
    }

    private async void Form1_Load(object? sender, EventArgs e)
    {
        _windowsSystemComPorts = await WindowsSystemComPorts.BuildAsync();
        _windowsSystemComPorts.StartWatchingPorts();
        _windowsSystemComPorts.SystemComPortAdded += _windowsSystemComPorts_SystemComPortAdded;
        _windowsSystemComPorts.SystemComPortRemoved += _windowsSystemComPorts_SystemComPortRemoved;
    }

    private void _windowsSystemComPorts_SystemComPortAdded(object? sender, ComPortEventArgs e)
    {
        Debug.WriteLine($"Connected: {e.PortName}");
    }

    private void _windowsSystemComPorts_SystemComPortRemoved(object? sender, ComPortEventArgs e)
    {
        Debug.WriteLine($"Disconnected: {e.PortName}");
    }

}
