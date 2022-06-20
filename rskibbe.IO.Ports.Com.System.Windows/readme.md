# Description
Easily detecting Windows COM port additions and removals with a few lines of code. [rskibbe.IO.Ports.Com](https://www.nuget.org/packages/rskibbe.IO.Ports.Com).

## Getting started
- **Import** the WindowsSystemComPorts-**Class** **from** the rskibbe.IO.Ports.Com.System.Windows-**Namespace**. 

      // normal
      using rskibbe.IO.Ports.Com.System.Windows;

      // using an alias
      using ComPorts = rskibbe.IO.Ports.Com.System.Windows.WindowsSystemComPorts;

- **Create an instance** of the class with the provided factory method:

      // normal
      var comPorts = await WindowsSystemComPorts.BuildAsync();

      // using the namespace alias
      var comPorts = await ComPorts.BuildAsync();

> Keep an eye on the _SynchronizationContext_-Parameter of the `BuildAsync`-Method. It defaults to the current one!

- **Attach a handler** to one or both events to start listening for changes

      private void ComPort_Added(object sender, ComPortEventArgs e)
      {
        // access the changed portName "COM3"
        // e.PortName
        // access the changed portId "COM3" -> 3
        // e.PortId
      }
      comPorts.SystemComPortAdded += ComPort_Added;