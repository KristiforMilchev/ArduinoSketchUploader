using System.IO.Ports;
 
namespace ArduinoUploader.BootloaderProgrammers.ResetBehavior
{
    internal interface IResetBehavior
    {
        SerialPort Reset(SerialPort serialPort, SerialPortConfig config);
    }
}