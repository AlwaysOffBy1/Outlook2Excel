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
                string filter = $"[UnRead]=true AND [ReceivedTime] >= '{DateTime.Now.AddDays(-5):g}'";
                var items = inbox.Items.Restrict(filter);

                //Each email returns a dictionary where KEY = property and VALUE = regex result
                List<Dictionary<string,string>> outputDictionaryList = new List<Dictionary<string,string>>();

                //Look up regex values in each email
                foreach(object item in items)
                {
                    if (item is not Outlook.MailItem mail) continue;
                    Outlook.MailItem mi = (Outlook.MailItem)item;
                    #pragma warning disable CS8604 // Possible null reference argument. AppSettings makes sure values are not null and quits if they are
                    Dictionary<string,string>? outputDictionary = disposableOutlook.GetValueFromEmail(mi, AppSettings.RegexMap, AppSettings.PrimaryKey);
                    if(outputDictionary != null) outputDictionaryList.Add(outputDictionary);
                    #pragma warning restore CS8604 // Possible null reference argument.
                }

                //Add each email to excel


                //TESTER
                int i = 1;
                foreach(Dictionary<string,string> outputDictionary in outputDictionaryList)
                {
                    Console.WriteLine($"Dictionary for email {i}");
                    i++;
                    foreach(var item in outputDictionary)
                    {
                        if (item.Key == "Body") continue;
                        Console.WriteLine("\t" +item.Key + ": " + item.Value);
                    }
                }

            }
            Console.WriteLine("Done.");
            Console.ReadLine();
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
