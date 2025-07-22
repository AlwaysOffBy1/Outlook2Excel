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
                List<Dictionary<string, string>> outputDictionaryList = GetDataFromOutlook();

                //Add each email to excel
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



        private List<Dictionary<string, string>> GetDataFromOutlook()
        {
            //Each email returns a dictionary where KEY = property and VALUE = regex result
            List<Dictionary<string, string>> outputDictionaryList = new List<Dictionary<string, string>>();

            Outlook2Excel.Core.AppLogger.Log.Info("Creating Outlook instance");
            using (DisposableOutlook disposableOutlook = new DisposableOutlook(AppSettings.Mailbox))
            {
                Outlook2Excel.Core.AppLogger.Log.Info("Outlook instance created");
                var recipient = disposableOutlook.Recipient;
                recipient.Resolve();
                if (!recipient.Resolved) StaticMethods.Quit("Could not access outlook", 201);

                Outlook.MAPIFolder? inbox = null;
                try
                {
                    inbox = disposableOutlook.Namespace.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderInbox);
                }
                catch (System.Exception ex)
                {
                    StaticMethods.Quit($"The mailbox {AppSettings.Mailbox} in appsettings.json is inaccessible to this PC. Please add the mailbox to Outlook and try again", 100);
                    return new List<Dictionary<string, string>>(); //need this to elimiate possible null reference of inbox warn
                }

                Outlook2Excel.Core.AppLogger.Log.Info("Sorting inbox...");
                string filter = $"[ReceivedTime] >= '{DateTime.Now.AddDays(0 - AppSettings.DaysToGoBack):g}'";
                var items = inbox.Items.Restrict(filter);

                //Look up regex values in each email
                foreach (object item in items)
                {
                    if (item is not Outlook.MailItem mail) continue;
                    Outlook.MailItem mi = (Outlook.MailItem)item;

                    //AppSettings makes sure values are not null and quits if they are
                    Dictionary<string, string>? outputDictionary = disposableOutlook.GetValueFromEmail(mi, AppSettings.RegexMap, AppSettings.PrimaryKey);
                    if(outputDictionary != null)
                    {
                        if (AppSettings.ImportDate) outputDictionary.Add("Date", DateTime.Now.ToString("MM/dd/yy hh:mm tt"));
                        outputDictionaryList.Add(outputDictionary);
                    }
                }
            }
            return outputDictionaryList;
        }


        //Inotifypropchanged

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    }
}
