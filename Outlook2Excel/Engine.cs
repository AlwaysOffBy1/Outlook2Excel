using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Timer = System.Timers.Timer;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Timers;
using System.Diagnostics;
using System.ComponentModel;

namespace Outlook2Excel.Core
{
    public class Engine : IDisposable, INotifyPropertyChanged
    {
        
        public string Progress;
        public DisposableExcel _disposableExcel;
        private Timer _timer;
        private string lastRan = "";

        public string LastRan
        {
            get => lastRan;
            set 
            {
                if (lastRan != value)
                {
                    lastRan = value;
                    OnPropertyChanged(nameof(LastRan));
                }
            }
        }

        public Engine() 
        {
            Progress = "Initializing";
            if (!AppSettings.GetSettings()) StaticMethods.Quit("Not all app settings were imported. Check app settings file.", 101);
            
            //This opens Excel, which should stay open
            Debug.WriteLine("Opening Excel...");
            _disposableExcel = new DisposableExcel(AppSettings.ExcelFilePath);

            _timer = new Timer(AppSettings.TimerInterval * 60 * 1000); //5 minutes is default
            _timer.AutoReset = true;
            _timer.Elapsed += TimerTicked;
            _timer.Enabled = true;
            _timer.Start();
            
            Progress = "Initialized";
        }

        private void TimerTicked(object sender, ElapsedEventArgs e)
        {
            RunNow();
            LastRan = $"Last ran - {DateTime.Now.ToString("MM/dd/yy - hh:mm tt")}";
        }

        //Public API
        public void RunNow() 
        {
            System.Diagnostics.Debug.WriteLine("Starting now...");
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
            System.Diagnostics.Debug.WriteLine("Finshed");
        }
        public void Pause() 
        { 
            _timer.Stop();
        }
        public void UnPause()
        {
            _timer.Start();
        }
        public void SetRunInterval(int intervalInMinutes)
        {
            _timer.Stop();
            if(intervalInMinutes <= 0) intervalInMinutes = AppSettings.TimerInterval;
            _timer.Interval = intervalInMinutes;
            _timer.Start();
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
            Outlook2Excel.Core.AppLogger.Log.Info("Creating Outlook instance");

            try
            {
                //DisposableOutlook handles all it's child COM objects upon disposal
                using (DisposableOutlook disposableOutlook = new DisposableOutlook(AppSettings.Mailbox,AppSettings.SubFolder, inboxSortFilter, AppSettings.RegexMap, AppSettings.PrimaryKey))
                {
                    Outlook2Excel.Core.AppLogger.Log.Info("Outlook instance created");
                    try
                    {
                        return disposableOutlook.GetEmailListFromOutlookViaRegexLookup();
                    }
                    catch(Exception e)
                    {
                        Outlook2Excel.Core.AppLogger.Log.Error("Outlook could not get emails", e);
                        StaticMethods.Quit("Outlook could not get emails", 200);
                        return null;
                    }
                    
                }
            }
            catch(Exception e)
            {
                Outlook2Excel.Core.AppLogger.Log.Error("Outlook instance failed to create.", e);
                StaticMethods.Quit("Outlook instance failed to create.", 200);
                
                //This is inaccessible since StaticMethods.Quit closes the app, but requried.
                return null;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    }
}
