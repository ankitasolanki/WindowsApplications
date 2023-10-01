using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessExcels.Utility
{
    public interface ILogger
    {
        void LogInfo(string data);
        void LogDebug(string data);
        void LogError(string module, string method, Exception exception);
    }
}
