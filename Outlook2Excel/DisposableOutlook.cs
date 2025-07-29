using System;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
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
        private Outlook.Folder? _folder;
        private Outlook.Items? _items;
        private List<Outlook.MailItem> _mailItems;
        private Outlook.MailItem? _currentMailItem;
        private List<object?> COM_OBJECTS;
        public string InboxSortFilter{ get; set; }
        public Dictionary<string, string> RegexMap { get; set; }
        public string PrimaryKey {  get; set; }
        public string EmailAccount { get; set; }
        private bool _disposed = false;


        public DisposableOutlook(string fullFolderPath, string? inboxSortFilter, Dictionary<string,string> regexMap, string? primaryKey)
        {
            //if provided sort filter is blank, set to "look at all emails within the past x days" where x is in Appsettings.json
            InboxSortFilter = inboxSortFilter ?? $"[ReceivedTime] >= '{DateTime.Now.AddDays(0 - AppSettings.DaysToGoBack):g}'";

            
            COM_OBJECTS = new List<object?>();
            RegexMap = regexMap ?? new Dictionary<string,string>();
            PrimaryKey = primaryKey ?? "";

            //Initialize all COMs,
            //_outlookApp -> _namespace -> _recipient -> _folder -> _items (set to _mailItems)

            try
            {
                //OUTLOOK APP
                try{
                    _outlookApp = new Outlook.Application();}
                catch (Exception ex){
                    throw new Exception("Failed to create Outlook Application", ex);}
                COM_OBJECTS.Add(_outlookApp);

                //NAMESPACE
                try{
                    _namespace = _outlookApp.GetNamespace("MAPI");}
                catch (Exception ex){
                    throw new Exception("Failed to get Outlook Namespace", ex);}

                //RECIPIENT / FOLDER
                if (!_SetFolderFromFullPath(fullFolderPath))
                    StaticMethods.Quit("Unable to get folder. Quitting.", 200, null);

                //ITEMS - FILTERED
                try{
                    _items = _folder?.Items.Restrict(InboxSortFilter);}
                catch (Exception ex){
                    throw new Exception("Failed to apply filter to inbox items: " + InboxSortFilter, ex);}

                try{
                    if (_items?.Count != 0)
                        _FilterCOMObjectsToMailItems();
                    else
                        _currentMailItem = null;}
                catch (Exception ex){
                    throw new Exception("Failed to retrieve first mail item from filtered inbox", ex);}
            }
            catch (Exception ex)
            {
                Dispose();
                StaticMethods.Quit(ex.Message, 200, ex);
            }
            


        }

        private bool _SetFolderFromFullPath(string fullEmailPath)
        {
            if (string.IsNullOrWhiteSpace(fullEmailPath))
                StaticMethods.Quit("FullEmailPath is empty.", 200, null);

            //Normalize and split path parts
            string[] parts = fullEmailPath.Split('\\', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                StaticMethods.Quit("FullEmailPath is invalid. Must be like: \\\\mailbox\\Inbox\\Subfolder...", 201, null);

            EmailAccount = parts[0];
            string[] folderPath = parts.Skip(1).ToArray();

            //Resolve recipient
            _recipient = _namespace.CreateRecipient(EmailAccount);
            COM_OBJECTS.Add(_recipient);
            _recipient.Resolve();

            if (!_recipient.Resolved)
                StaticMethods.Quit($"Mailbox '{EmailAccount}' could not be resolved.", 202, null);

            //Get the store root
            Outlook.Folder inbox = (Outlook.Folder)_namespace.GetSharedDefaultFolder(_recipient, Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Folder current = (Outlook.Folder)inbox.Parent;

            // Traverse the folder path step-by-step
            foreach (string folderName in folderPath)
            {
                Outlook.Folder? next = null;

                for (int i = 1; i <= current.Folders.Count; i++)
                {
                    var sub = (Outlook.Folder)current.Folders[i];
                    Debug.WriteLine("Found " + current.Name + "\\" + sub.Name);

                    if (sub.Name.Trim().Equals(folderName.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        next = sub;
                        break;
                    }

                    DisposeObject(sub);
                }

                if (next == null & next.Name != "Inbox")
                    StaticMethods.Quit($"Folder '{folderName}' not found under '{current.Name}'.", 203, null);

                if (!ReferenceEquals(current, next))
                    DisposeObject(current);

                current = next;
            }

            _folder = current;
            AppLogger.Log.Info($"Folder set to: {_folder.FolderPath}");
            return true;
        }

        private bool _FilterCOMObjectsToMailItems()
        {
            if (_items == null)
                StaticMethods.Quit("Outlook mail items is null", 202, null);

            //Temporary list to hold valid mail items
            List<Outlook.MailItem> filtered = new();

            for (int i = 1; i <= _items.Count; i++)
            {
                object item = _items[i];

                if (item is Outlook.MailItem mail)
                    filtered.Add(mail);
                else
                    DisposeObject(item);
            }

            //Release the original collection
            DisposeObject(_items);

            _mailItems = filtered;
            _currentMailItem = _mailItems[0];
            return true;
        }        
        public List<Dictionary<string, string>>? GetEmailListFromOutlookViaRegexLookup()
        {
            List<Dictionary<string,string>> outputDictionaryList = new List<Dictionary<string,string>>();
            //_mailItems could not be properly populated because it does not exist
            if (_mailItems == null) return null;
            //_currentItem is null because _mailItems exists, but is empty
            if(_currentMailItem == null) return outputDictionaryList;
            for (int i = 0; i < _mailItems.Count; i++)
            {
                if (_mailItems[i] is not Outlook.MailItem mail) continue;
                _currentMailItem = _mailItems[i];
                Debug.WriteLine(_currentMailItem.Subject);

                Dictionary<string, string>? outputDictionary = GetValueFromEmail();
                //If outlook found a matching email and got the regex results
                if (outputDictionary != null)
                {
                     outputDictionaryList.Add(outputDictionary);
                }
            }
            string sortBy = AppSettings.OrganizeBy;
            if (sortBy == "PrimaryKey") sortBy = AppSettings.PrimaryKey;
            var output = outputDictionaryList
                .DistinctBy(d => d[PrimaryKey])                                // remove duplicates by PrimaryKey value
                .OrderBy(d => d.TryGetValue(sortBy, out var val) ? val : "") // order by Date
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

        #region Disposals
        //Disposals
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
        protected void DisposeObjects(List<object?> os)
        {
            for(int i = 0; i < os.Count; i++)
            {
                DisposeObject(os[i]);
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
        #endregion
    }

}
