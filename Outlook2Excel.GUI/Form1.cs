using System.ComponentModel;
using System.Timers;
using Outlook2Excel.Core;
using Timer = System.Timers.Timer;

namespace Outlook2Excel.GUI
{
    public partial class Form1 : Form
    {
        private NotifyIcon _trayIcon;
        private ContextMenuStrip _trayMenu;
        private Outlook2Excel.Core.Engine? _engine;
        private ToolStripMenuItem _pauseItem;
        private ToolStripMenuItem _unpauseItem;
        private ToolStripMenuItem _lastRanItem;
        private Timer _timer;

        public string LastRan;

        public Form1()
        {
            InitializeComponent();

            Outlook2Excel.Core.AppLogger.Log.Info("Initializing...");

            //Timer to run the engine every x minutes
            _timer = new Timer(AppSettings.TimerInterval * 60 * 1000); //5 minutes is default
            _timer.AutoReset = true;
            _timer.Elapsed += TimerTicked;
            _timer.Enabled = true;

            _lastRanItem = new ToolStripMenuItem("Last ran - ");
            _lastRanItem.Enabled = false;
            _pauseItem = new ToolStripMenuItem("Pause", null, OnPause);
            _unpauseItem = new ToolStripMenuItem("Unpause", null, OnUnPause) { Visible = false };

            _trayMenu = new ContextMenuStrip();
            _trayMenu.Items.Add(_lastRanItem);
            _trayMenu.Items.Add("Run Now", null, OnRunNow);
            _trayMenu.Items.Add(_pauseItem);
            _trayMenu.Items.Add(_unpauseItem);
            _trayMenu.Items.Add("Exit", null, OnExit);

            

            _trayIcon = new NotifyIcon
            {
                Icon = new Icon("Outlook2Excel.ico"), //TODO change to icon
                ContextMenuStrip = _trayMenu,
                Visible = true,
                Text = "Outlook2Excel"
            };

            

            _timer.Start();

            //Hide the main window
            this.Shown += UserformLoaded;
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            Outlook2Excel.Core.AppLogger.Log.Info("Tray initialized sucessfully");
        }

        private void UserformLoaded(object? sender, EventArgs e)
        {
            _engine = new Outlook2Excel.Core.Engine();
            RunEngine();
        }
        

        #region Timer
        private void TimerTicked(object? sender, ElapsedEventArgs e) => RunEngine();
        public void Pause() => _timer.Stop();
        public void UnPause() => _timer.Start();
        public void SetRunInterval(int intervalInMinutes)
        {
            _timer.Stop();
            if (intervalInMinutes <= 0) intervalInMinutes = AppSettings.TimerInterval;
            _timer.Interval = intervalInMinutes;
            _timer.Start();
        }
        #endregion

        private void RunEngine()
        {
            if (_engine.IsRunning)
            {
                Outlook2Excel.Core.AppLogger.Log.Info("User initiated run while program already running");
                return;
            }
            Outlook2Excel.Core.AppLogger.Log.Info("Running...");
            LastRan = DateTime.Now.ToString("MM/dd - hh:mm tt");
            _engine.RunNow();
            Outlook2Excel.Core.AppLogger.Log.Info("Finished running");
            Invoke(() => _lastRanItem.Text = LastRan);

            //test
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #region Button Click Events
        private void OnRunNow(object? sender, EventArgs e) => RunEngine();
        private void OnPause(object? sender, EventArgs e)
        {
            Pause();
            Outlook2Excel.Core.AppLogger.Log.Info("Paused");
            _pauseItem.Visible = false;
            _unpauseItem.Visible = true;
        }
        private void OnUnPause(object? sender, EventArgs e) 
        {
            UnPause();
            Outlook2Excel.Core.AppLogger.Log.Info("Unpaused");
            _pauseItem.Visible = true;
            _unpauseItem.Visible = false;
        }
        private void OnExit(object? sender, EventArgs e)
        {
            Outlook2Excel.Core.AppLogger.Log.Info("Exiting");
            _engine.Dispose();
            _trayIcon.Visible = false;
            Application.Exit();
        }
        #endregion
    }
}
