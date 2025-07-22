using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace Outlook2Excel.Core
{
    public static class AppLogger
    {
        public static readonly ILog Log = LogManager.GetLogger("AppLogger");
        
    }
    
}
