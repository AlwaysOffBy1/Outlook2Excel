using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

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
