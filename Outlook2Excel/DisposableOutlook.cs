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
using Outlook2Excel.Core;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace Outlook2Excel
{
    public class DisposableOutlook : IDisposable
    {
        private Outlook.Application? _outlookApp;
        private Outlook.NameSpace? _namespace;
        private Outlook.Recipient? _recipient;
        private Outlook.MAPIFolder? _inbox;
        private Outlook.Items? _items;
        private Outlook.MailItem? _currentMailItem;
        private List<object?> COM_OBJECTS;
        public string InboxSortFilter{ get; set; }
        public Dictionary<string, string> RegexMap { get; set; }
        public string PrimaryKey {  get; set; }

        private bool _disposed = false;

        #pragma warning disable CS8618 //Program doesn't continue if any field is null, so the warn is not required
        public DisposableOutlook(string mailbox, string? inboxSortFilter, Dictionary<string,string> regexMap, string? primaryKey)
        #pragma warning restore CS8618 //Program doesn't continue if any field is null, so the warn is not required
        {
            //if provided sort filter is blank, set to "look at all emails within the past x days" where x is in Appsettings.json
            InboxSortFilter = inboxSortFilter ?? $"[ReceivedTime] >= '{DateTime.Now.AddDays(0 - AppSettings.DaysToGoBack):g}'";
            COM_OBJECTS = new List<object?>();
            RegexMap = regexMap ?? new Dictionary<string,string>();
            PrimaryKey = primaryKey ?? "";

            //Each COM object can fail to initialize for a different reason
            //This is the best I could come up with to ensure good error handling
            //changed spacing because it looks bloated and making it short makes me feel better
            try{
                
                try{
                    _outlookApp = new Outlook.Application();}
                catch (Exception ex){
                    throw new Exception("Failed to create Outlook Application", ex);}
                
                try{
                    _namespace = _outlookApp.GetNamespace("MAPI");}
                catch (Exception ex){
                    throw new Exception("Failed to get Outlook Namespace", ex);}
                
                try{
                    _recipient = _namespace.CreateRecipient(mailbox); 
                    _recipient.Resolve();}
                catch (Exception ex){
                    throw new Exception("Failed to create recipient for mailbox: " + mailbox, ex);}

                try{
                    _inbox = _namespace.GetSharedDefaultFolder(_recipient, Outlook.OlDefaultFolders.olFolderInbox);}
                catch (Exception ex){
                    throw new Exception($"The mailbox \"{AppSettings.Mailbox}\" in appsettings.json is inaccessible to this PC.\n\n Please fix the name to an accessible mailbox and try again.", ex);}

                try{
                    _items = _inbox.Items.Restrict(InboxSortFilter);}
                catch (Exception ex){
                    throw new Exception("Failed to apply filter to inbox items: " + InboxSortFilter, ex);}

                try{
                    _currentMailItem = _items[0];}
                catch (Exception ex){
                    throw new Exception("Failed to retrieve first mail item from filtered inbox", ex);}


                if (!_recipient.Resolved)
                    throw new Exception("Recipient could not be resolved");

                //Add all objects to a list to make disposing less lines of code
                COM_OBJECTS.AddRange(new List<object?>
                {
                    _outlookApp, 
                    _namespace, 
                    _recipient, 
                    _inbox, 
                    _currentMailItem
                });

                Outlook2Excel.Core.AppLogger.Log.Info("Outlook and all dependents created successfully.");
            }
            catch (Exception ex)
            {
                Dispose();
                StaticMethods.Quit(ex.ToString(), 200);
            }


        }      

        public List<Dictionary<string, string>>? GetEmailListFromOutlookViaRegexLookup()
        {
            List<Dictionary<string,string>> outputDictionaryList = new List<Dictionary<string,string>>();
            if (_currentMailItem == null || _items == null) return null;
            for (int i = 0; i < _items.Count; i++)
            {
                if (_items[i] is not Outlook.MailItem mail) continue;
                _currentMailItem = _items[i];

                Dictionary<string, string>? outputDictionary = GetValueFromEmail();
                if (outputDictionary != null)
                {
                    outputDictionaryList.Add(outputDictionary);
                }
            }
            return outputDictionaryList;
        }
        private Dictionary<string,string>? GetValueFromEmail()
        {
            Dictionary<string,string> output = new Dictionary<string,string>();
            if (_currentMailItem == null) throw new Exception("Current mail item is null. Outlook can not access a null mail item");
            string message = _currentMailItem.Subject + "\n\n" + _currentMailItem.Body;
            
            //Before doing anything, if we have a primary key, check the email for it first
            if (PrimaryKey != "")
            {
                output.Add(PrimaryKey, EmailRegexSearchFor(message, RegexMap[PrimaryKey]) ?? "");
                if (output[PrimaryKey] == "" || output[PrimaryKey] == null) return null;
            }

            if(AppSettings.ImportDate) output.Add("Date", _currentMailItem.ReceivedTime.ToString("MM/dd/yyyy hh:mm tt"));
            output.Add("Subject", _currentMailItem.Subject);
            output.Add("Body", _currentMailItem.Body);

            //Loop through regex map to search for properties
            foreach (var pair in RegexMap) 
            {
                if (pair.Key == PrimaryKey) continue;
                string? foundVal = EmailRegexSearchFor(message, pair.Value);
                if (foundVal != null) output.Add(pair.Key, foundVal);
            }
            return output;
        }
        private string? EmailRegexSearchFor(string message, string pattern)
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

            //Release unmanaged COM objects
            for (int i = 0; i < COM_OBJECTS.Count; i++)
            {
                object? o = COM_OBJECTS[i];
                if (o != null)
                {
                    Marshal.ReleaseComObject(o);
                    COM_OBJECTS[i] = null; 
                }
            }
            _disposed = true;
        }

        ~DisposableOutlook()
        {
            Dispose(false);
        }
    }

}
