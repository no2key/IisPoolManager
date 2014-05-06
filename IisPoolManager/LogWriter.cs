using System;
using System.IO;

namespace IisPoolManager
{
    internal sealed class LogWriter
    {
        public static void WriteLog(DateTime time, string message)
        {
            var name = "Log_" + DateTime.Today.ToString().Substring(0, 10) + ".txt";
            using (var writer = new StreamWriter(name, true))
            {
                writer.WriteLine(time + "     " + message);
                writer.Close();
            }
        }
    }
}