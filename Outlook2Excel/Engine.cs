
using System.Timers;
using System.Diagnostics;
using System.ComponentModel;

namespace Outlook2Excel.Core
{
    public class Engine : IDisposable
    {
        
        public string Progress;
        public DisposableExcel _disposableExcel;
        private string lastRan = "";
        private bool isRunning;

        public bool IsRunning = false;

        public Engine() 
        {
            AppLogger.Log.Info("Initializing engine");
            if (!AppSettings.GetSettings()) StaticMethods.Quit("Not all app settings were imported. Check app settings file.", 101, null);

            //This opens Excel, which should stay open
            AppLogger.Log.Info("Opening Excel");
            _disposableExcel = new DisposableExcel(AppSettings.ExcelFilePath);
        }

        

        //Public API
        public void RunNow() 
        {
            System.Diagnostics.Debug.WriteLine("Starting now...");
            IsRunning = true;
            //Prevent UI lockup with Task.Run
            Task.Run(() =>
            {
                //Returns a list (each email) of dictionary<string,string> (The lookup key and lookup result per email)
                List<Dictionary<string, string>>? outputDictionaryList = GetDataFromOutlook();

                //Add each email to excel if its not null
                if (outputDictionaryList == null) 
                    AppLogger.Log.Warn("OUTLOOK FAILED TO GET DATA");
                else
                    _disposableExcel.AddData(outputDictionaryList, AppSettings.PrimaryKey);
            });
            IsRunning = false;
            System.Diagnostics.Debug.WriteLine("Finshed");
        }
        
        public string GetStatus() { return Progress; }

        public void Dispose()
        {
            _disposableExcel?.SaveAndClose();
            _disposableExcel?.Dispose();
        }
        public List<Dictionary<string, string>>? GetDataFromOutlook()
        {
            //Each email returns a dictionary where KEY = property and VALUE = regex result
            List<Dictionary<string, string>> outputDictionaryList = new List<Dictionary<string, string>>();
            string inboxSortFilter = $"[ReceivedTime] >= '{DateTime.Now.AddDays(0 - AppSettings.DaysToGoBack):g}'";
            if (!string.IsNullOrEmpty(AppSettings.SubjectFilter)) inboxSortFilter += $" AND [Subject] LIKE '%{AppSettings.SubjectFilter}%'";
            Outlook2Excel.Core.AppLogger.Log.Info("Creating Outlook instance");

            try
            {
                //DisposableOutlook handles all it's child COM objects upon disposal
                using (DisposableOutlook disposableOutlook = new DisposableOutlook(AppSettings.Mailbox, AppSettings.SubFolder, inboxSortFilter, AppSettings.RegexMap, AppSettings.PrimaryKey))
                {
                    Outlook2Excel.Core.AppLogger.Log.Info("Outlook instance created");
                    try
                    {
                        return disposableOutlook.GetEmailListFromOutlookViaRegexLookup();
                    }
                    catch(Exception ex)
                    {
                        StaticMethods.Quit("Outlook could not get emails", 200, ex);
                        //This is inaccessible since StaticMethods.Quit closes the app, but requried.
                        return null;
                    }
                    
                }
            }
            catch(Exception ex)
            {
                StaticMethods.Quit("Outlook instance failed to create.", 200, ex);
                //This is inaccessible since StaticMethods.Quit closes the app, but requried.
                return null;
            }
        }
    }
}
