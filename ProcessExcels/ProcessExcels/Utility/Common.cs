using System.IO;

namespace ProcessExcels.Utility
{
    static class Common
    {
        public static string GetApplicationPath()
        {
            return Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
        }

        public static string GetLogFilePath()
        {
            return Path.Combine(GetApplicationPath(), "ProcessExcels.txt");
        }

        public static string GetErrorFilePath()
        {
            return Path.Combine(GetApplicationPath(), "ProcessExcelsError.txt");
        }
    }
}
