namespace rskibbe.IO.Ports.Com.System.Windows.Tests;

public class WindowsSystemComPortsTests
{
    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public void TestList()
    {
        // only if windows is currently having 2 ports..
        // not going to go this way further right now..
        // just a little helper for me during dev
        var comPorts = new WindowsSystemComPorts();
        comPorts.Initialized += (s, e) =>
        {
            var ports = comPorts.ListUsedPortNamesAsync().Result;
            Assert.That(ports.Count(), Is.EqualTo(2));
        };
    }

}