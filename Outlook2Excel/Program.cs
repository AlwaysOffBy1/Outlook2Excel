using System;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Outlook;

namespace Outlook2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting EmailOrderProcessor...");
            Console.WriteLine("Importing App settings");

            if (!AppSettings.GetSettings()) Quit("Not all app settings were imported. Check app settings file.");

            using (DisposableOutlook disposableOutlook = new DisposableOutlook(AppSettings.Mailbox))
            {
                var recipient = disposableOutlook.Recipient;
                recipient.Resolve();
                if (!recipient.Resolved) Quit("Could not access outlook");
                Console.WriteLine($"Resolved: {recipient.Name}");

                Console.WriteLine("Reading inbox...");
                var inbox = disposableOutlook.Namespace.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderInbox);
                var filter = $"[UnRead]=true AND [ReceivedTime] >= '{DateTime.Now.AddDays(-5):g}'";
                var items = inbox.Items.Restrict(filter);

            }
            
            

            //foreach (object item in items)
            //{
            //    if (item is Outlook.MailItem mail)
            //    {
            //        string subject = mail.Subject;
            //        string sender = mail.SenderName;
            //        DateTime received = mail.ReceivedTime;

            //        string orderNumber = ExtractOrderNumber(subject);
            //        if (!string.IsNullOrEmpty(orderNumber))
            //        {
            //            Console.WriteLine($"Order #{orderNumber} from {sender} at {received}");
            //            // TODO: Write to Excel here
            //        }

            //        mail.UnRead = false; // mark as read
            //        mail.Save();
            //    }
            //}

            Console.WriteLine("Done.");
        }

        static string ExtractOrderNumber(string subject)
        {
            var match = Regex.Match(subject, @"Order\s*#?(\d+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : null;
        }

        static void Quit(string reason)
        {
            Console.WriteLine(reason);
            Environment.Exit(0);
        }
    }
}
