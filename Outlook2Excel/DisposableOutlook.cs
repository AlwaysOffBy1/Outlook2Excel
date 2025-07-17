using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;

namespace Outlook2Excel
{
    public class DisposableOutlook : IDisposable
    {
        private Outlook.Application _outlookApp;
        private Outlook.NameSpace _namespace;
        private Outlook.Recipient _recipient;

        private bool _disposed = false;

        public DisposableOutlook(string mailbox)
        {
            _outlookApp = new Outlook.Application();
            _namespace = _outlookApp.GetNamespace("MAPI");
            _recipient = _namespace.CreateRecipient(mailbox);
        }
        public Outlook.Application App => _outlookApp;
        public Outlook.NameSpace Namespace => _namespace;
        public Outlook.Recipient Recipient => _recipient;
        
        public Dictionary<string,string>? GetValueFromEmail(Outlook.MailItem item, Dictionary<string,string> regexMap, string primaryKey)
        {
            Dictionary<string,string> output = new Dictionary<string,string>();

            string message = item.Subject + "\n\n" + item.Body;

            //Before doing anything, if we have a primary key, check the email for it first
            if (primaryKey != "")
            {
                output.Add(primaryKey, SearchFor(message, regexMap[primaryKey]) ?? "");
                if (output[primaryKey] == "" || output[primaryKey] == null) return null;
            }

            output.Add("Email Date", item.ReceivedTime.ToString("MM/dd/yyyy hh:mm tt"));
            output.Add("Subject", item.Subject);
            output.Add("Body", item.Body);

            //Loop through regex map to search for properties
            foreach (var pair in regexMap) 
            {
                if (pair.Key == primaryKey) continue;
                string? foundVal = SearchFor(message, pair.Value);
                if (foundVal != null) output.Add(pair.Key, foundVal);
            }
            return output;
        }
        private string? SearchFor(string message, string pattern)
        {
            //Instead of looping through all properties, separated this part so PrimaryKey can be searched for first
            var match = Regex.Match(message, pattern);
            if (match.Success && 
                match.Groups.Count > 0) return match.Groups[1].Value.Trim();
            else return null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }  

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // Release managed resources if any (none in this case)
            }

            // Release unmanaged COM objects
            if (_namespace != null)
            {
                Marshal.ReleaseComObject(_namespace);
                _namespace = null;
            }

            if (_outlookApp != null)
            {
                Marshal.ReleaseComObject(_outlookApp);
                _outlookApp = null;
            }

            _disposed = true;
        }

        ~DisposableOutlook()
        {
            Dispose(false);
        }
    }

}
