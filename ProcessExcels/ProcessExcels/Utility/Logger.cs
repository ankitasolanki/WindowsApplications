using System;
using System.IO;

namespace ProcessExcels.Utility
{
    public class Logger : ILogger
    {
        public void LogDebug(string data)
        {
            writelog("DEBUG", data);
        }

        public void LogError(string module, string method, Exception exception)
        {
            writeError(module, method,exception);
        }

        public void LogInfo(string data)
        {
            writelog("INFO", data);
        }

        private void writelog(string type, string data)
        {
            try
            {
                string logfilePath = Common.GetLogFilePath();

                using (StreamWriter writer = File.AppendText(logfilePath))
                {
                    writer.WriteLine($"[{type}]:{DateTime.Now}:{data}");
                }
            }
            catch
            {

            }
        }

        private void writeError(string module, string method, Exception exception)
        {
            try
            {
                string logfilePath = Common.GetErrorFilePath();
                using (StreamWriter writer = File.AppendText(logfilePath))
                {
                    writer.WriteLine($"[ERROR]:{DateTime.Now}:{module}:{method}: Error: {exception.Message}\r\n{exception.StackTrace}");
                }
            }
            catch
            {

            }
        }
    }
}
