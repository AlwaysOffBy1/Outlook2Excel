using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Outlook2Excel.Core
{
    public static class StaticMethods
    {
        public static void Quit(string reason, int errorCode)
        {
            Outlook2Excel.Core.AppLogger.Log.Error(reason);
            Console.WriteLine(reason);
            Environment.Exit(errorCode);
        }
    }
}
