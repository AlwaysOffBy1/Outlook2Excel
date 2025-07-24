using System;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Outlook2Excel.Core;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;

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
        public DisposableOutlook(string mailbox, string? subFolder, string? inboxSortFilter, Dictionary<string,string> regexMap, string? primaryKey)
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

                if (!_recipient.Resolved)
                    throw new Exception("Recipient could not be resolved");

                try{
                    _inbox = SetInboxSubfolder(mailbox, subFolder);
                }
                catch (Exception ex){
                    throw new Exception($"The mailbox \"{mailbox}{(subFolder == "" || subFolder == null ? $"": $"/Inbox/{subFolder}")}\" in appsettings.json are inaccessible to this PC.\n\n Please fix the name to an accessible mailbox and try again.", ex);}

                try{
                    _items = _inbox.Items.Restrict(InboxSortFilter);}
                catch (Exception ex){
                    throw new Exception("Failed to apply filter to inbox items: " + InboxSortFilter, ex);}

                try{
                    if (_items.Count != 0)
                        _currentMailItem = _items[1];
                    else
                        _currentMailItem = null;
                }
                catch (Exception ex){
                    throw new Exception("Failed to retrieve first mail item from filtered inbox", ex);}

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
                StaticMethods.Quit("Unable to boot up Outlook", 200, ex);
            }


        }      
        private Outlook.MAPIFolder SetInboxSubfolder(string mailbox, string? subfolder)
        {
            //Double-check nulls
            if (_namespace == null) 
                throw new Exception("Namespace is null while trying to access folder");
            if (_recipient == null)
                throw new Exception("Recipient is null while trying to access folder");

            if (string.IsNullOrEmpty(subfolder)) return _namespace.GetSharedDefaultFolder(_recipient, Outlook.OlDefaultFolders.olFolderInbox);

            if (!_recipient.Resolved)
                throw new Exception("Recipient not resolved while trying to access folder");

            //Get root dir
            var root = _namespace.GetSharedDefaultFolder(_recipient, Outlook.OlDefaultFolders.olFolderInbox).Parent as Outlook.MAPIFolder
                        ?? throw new Exception("Unable to get mailbox root");
            var current = root;

            bool found = false;

            //Loop through folders and subfolders for the proper name in appconfig.json
            foreach (var part in subfolder.Split('/'))
            {
                Debug.WriteLine("Part: " + part);
                _ = current.Folders.Count; //force load
                
                string s = current.Name;

                for (int i = 1; i <= current.Folders.Count; i++)
                {
                    var folder = current.Folders[i];
                    Debug.WriteLine("    Folder: " + folder.Name);

                    if (folder.Name.Equals(part, StringComparison.OrdinalIgnoreCase))
                    {
                        current = folder;
                        found = true;
                        break;
                    }
                    DisposeObject(folder);
                    
                }   
            }
            DisposeObject (root);
            DisposeObject (subfolder);
            DisposeObject (current);


            if (!found) return _namespace.GetSharedDefaultFolder(_recipient, Outlook.OlDefaultFolders.olFolderInbox); ;
            return current;


        }

        public List<Dictionary<string, string>>? GetEmailListFromOutlookViaRegexLookup()
        {
            List<Dictionary<string,string>> outputDictionaryList = new List<Dictionary<string,string>>();
            //_items could not be properly populated because it does not exist
            if (_items == null) return null;
            //_currentItem is null because items exists, but is empty
            if(_currentMailItem == null) return outputDictionaryList;
            for (int i = 1; i < _items.Count; i++)
            {
                if (_items[i] is not Outlook.MailItem mail) continue;
                _currentMailItem = _items[i];

                Dictionary<string, string>? outputDictionary = GetValueFromEmail();
                //If outlook found a matching email and got the regex results
                if (outputDictionary != null)
                {
                     outputDictionaryList.Add(outputDictionary);
                }
            }

            var output = outputDictionaryList
                .DistinctBy(d => d[PrimaryKey])                                // remove duplicates by PrimaryKey value
                .OrderBy(d => d.TryGetValue(d[AppSettings.OrganizeBy], out var val) ? val : "") // order by Date
                .ToList();

            //testing
            foreach (var el in output)
            {
                Debug.WriteLine("Primary Key = " + el[PrimaryKey]);
            }

            return output;
        }
        private Dictionary<string,string>? GetValueFromEmail()
        {
            Dictionary<string,string> output = new Dictionary<string,string>();
            if (_currentMailItem == null) throw new Exception("Current mail item is null. Outlook can not access a null mail item");
            string message = _currentMailItem.Subject + "\n\n" + _currentMailItem.Body;

            output.Add("Subject", _currentMailItem.Subject);
            output.Add("Body", _currentMailItem.Body);
            output.Add("EmailDate", _currentMailItem.ReceivedTime.ToString("MM/dd/yyyy hh:mm tt"));

            //Before doing anything, if we have a primary key, check the email for it first
            if (PrimaryKey != "")
            {
                output.Add(PrimaryKey, EmailRegexSearchFor(message, RegexMap[PrimaryKey]) ?? "");
                if (output[PrimaryKey] == "" || output[PrimaryKey] == null) return null;
            }

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

        protected void DisposeObject(object? o)
        {
            if (o != null)
            {
                Marshal.ReleaseComObject(o);
                o = null;
            }
        }
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            //Release unmanaged COM objects
            for (int i = 0; i < COM_OBJECTS.Count; i++)
            {
                DisposeObject(COM_OBJECTS[i]);
            }
            _disposed = true;
        }

        ~DisposableOutlook()
        {
            Dispose(false);
        }
    }

}
