using System;
using System.IO;
using System.Text;

namespace NovemberFirstPlugin
{
    public class Logger
    {
        public static void logMessage(string msg)
        {
            try
            {
                FileStream fileStream = File.Open("c:/Uniconta/NovemberFirstPlugin.log", FileMode.Append);
                byte[] bytes = Encoding.ASCII.GetBytes(string.Format("{0} {1}\r\n", (object)DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"), (object)msg));
                fileStream.Write(bytes, 0, bytes.Length);
                fileStream.Close();
            }
            catch (Exception ex)
            {
            }
        }
    }
}
