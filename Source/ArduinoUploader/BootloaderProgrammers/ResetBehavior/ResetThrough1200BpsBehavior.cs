using System.IO.Ports;

namespace ArduinoUploader.BootloaderProgrammers.ResetBehavior
{
    internal class ResetThrough1200BpsBehavior : IResetBehavior
    {
        private static IArduinoUploaderLogger Logger => ArduinoSketchUploader.Logger;

        public SerialPort Reset(SerialPort serialPort, SerialPortConfig config)
        {
            const int timeoutVirtualPortDiscovery = 10000;
            const int virtualPortDiscoveryInterval = 100;
            Logger?.Info("Issuing forced 1200bps reset...");
            var currentPortName = serialPort.PortName;
            var originalPorts = SerialPort.GetPortNames();

            // Close port ...
            serialPort.Close();

            // And now open port at 1200 bps
            serialPort = new SerialPort(currentPortName, 1200)
            {
              //  Handshake = Handshake.DtrRts
            };
            serialPort.Open();

            // Close and wait for a new virtual COM port to appear ...
            serialPort.Close();

            var newPort = WaitHelper.WaitFor(timeoutVirtualPortDiscovery, virtualPortDiscoveryInterval,
                () => SerialPort.GetPortNames().Except(originalPorts).SingleOrDefault(),
                (i, item, interval) =>
                    item == null
                        ? $"T+{i * interval} - Port not found"
                        : $"T+{i * interval} - Port found: {item}");

            if (newPort == null)
                throw new ArduinoUploaderException(
                    $"No (unambiguous) virtual COM port detected (after {timeoutVirtualPortDiscovery}ms).");

            return new SerialPort()
            {
                BaudRate = config.BaudRate,
                PortName = newPort,
                DataBits = 8,
                Parity = Parity.None,
                StopBits = StopBits.One,
              //  Handshake = Handshake.DtrRts
            };
        }
    }
}