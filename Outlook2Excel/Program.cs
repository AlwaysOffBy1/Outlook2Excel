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

            if (!AppSettings.GetSettings()) Quit("Not all app settings were imported. Check app settings file.", 101);

            //Each email returns a dictionary where KEY = property and VALUE = regex result
            List<Dictionary<string, string>> outputDictionaryList = new List<Dictionary<string, string>>();

            using (DisposableOutlook disposableOutlook = new DisposableOutlook(AppSettings.Mailbox))
            {
                var recipient = disposableOutlook.Recipient;
                recipient.Resolve();
                if (!recipient.Resolved) Quit("Could not access outlook", 201);
                Console.WriteLine($"Resolved: {recipient.Name}");

                Console.WriteLine("Reading inbox...");
                Outlook.MAPIFolder? inbox = null;
                try
                {
                    inbox = disposableOutlook.Namespace.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderInbox);
                }
                catch (System.Exception ex)
                {
                    Quit($"The mailbox {AppSettings.Mailbox} in appsettings.json is inaccessible to this PC. Please add the mailbox to Outlook and try again", 100);
                    return; //need this to elimiate possible null reference of inbox warn
                }
                
                Console.WriteLine("Sorting inbox...");
                string filter = $"[UnRead]=true AND [ReceivedTime] >= '{DateTime.Now.AddDays(0-AppSettings.DaysToGoBack):g}'";
                var items = inbox.Items.Restrict(filter);

                //Look up regex values in each email
                foreach(object item in items)
                {
                    if (item is not Outlook.MailItem mail) continue;
                    Outlook.MailItem mi = (Outlook.MailItem)item;
                    Console.WriteLine($"Reading inbox email - {mi.Subject}");
                    //AppSettings makes sure values are not null and quits if they are
                    Dictionary<string,string>? outputDictionary = disposableOutlook.GetValueFromEmail(mi, AppSettings.RegexMap, AppSettings.PrimaryKey);
                    
                    //Just a difficult way of showing the user if their primary key was found, or if they dont have one, that an email was found
                    if (outputDictionary == null)
                    {
                        Console.WriteLine($"No {(AppSettings.PrimaryKey != "" ? "email":AppSettings.PrimaryKey)} found");
                    }
                    else
                    {
                        Console.WriteLine($"Found {(AppSettings.PrimaryKey != "" ? $"key {outputDictionary[AppSettings.PrimaryKey]}" : "email")}");
                        outputDictionaryList.Add(outputDictionary);
                    }
                }
            }

            //Add each email to excel
            using (DisposableExcel disposableExcel = new DisposableExcel())
            {
                disposableExcel.AddData(outputDictionaryList);
            }

            Console.WriteLine("Done.");
            Console.ReadLine();
        }



        public static void Quit(string reason, int errorCode)
        {
            Console.WriteLine(reason);
            Environment.Exit(errorCode);
        }
    }
}
