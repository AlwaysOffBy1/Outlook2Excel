using System;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace Outlook2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting EmailOrderProcessor...");
            Console.WriteLine("Importing App settings");

            var config = new ConfigurationBuilder()
                 .SetBasePath(Directory.GetCurrentDirectory())
                 .AddJsonFile("appsettings.json")
                 .Build();

            string? mailbox = config["Mailbox"];
            if(string.IsNullOrEmpty(mailbox))
            {
                Console.WriteLine("Mailbox configuration is missing.");
                return;
            }

            var outlookApp = new Outlook.Application();
            var ns = outlookApp.GetNamespace("MAPI");

            var recipient = ns.CreateRecipient(mailbox);
            recipient.Resolve();

            if (!recipient.Resolved)
            {
                Console.WriteLine("Failed to resolve shared mailbox.");
                return;
            }

            var inbox = ns.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderInbox);
            var items = inbox.Items.Restrict("[UnRead]=true");

            foreach (object item in items)
            {
                if (item is Outlook.MailItem mail)
                {
                    string subject = mail.Subject;
                    string sender = mail.SenderName;
                    DateTime received = mail.ReceivedTime;

                    string orderNumber = ExtractOrderNumber(subject);
                    if (!string.IsNullOrEmpty(orderNumber))
                    {
                        Console.WriteLine($"Order #{orderNumber} from {sender} at {received}");
                        // TODO: Write to Excel here
                    }

                    mail.UnRead = false; // mark as read
                    mail.Save();
                }
            }

            Console.WriteLine("Done.");
        }

        static string ExtractOrderNumber(string subject)
        {
            var match = Regex.Match(subject, @"Order\s*#?(\d+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : null;
        }
    }
}
