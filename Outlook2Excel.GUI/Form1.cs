using System.ComponentModel;
using Outlook2Excel.Core;

namespace Outlook2Excel.GUI
{
    public partial class Form1 : Form
    {
        private NotifyIcon _trayIcon;
        private ContextMenuStrip _trayMenu;
        private Outlook2Excel.Core.Engine _engine;
        private ToolStripMenuItem _pauseItem;
        private ToolStripMenuItem _unpauseItem;
        private ToolStripMenuItem _lastRanItem;

        public Form1()
        {
            InitializeComponent();

            _lastRanItem = new ToolStripMenuItem("Last ran - ");
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
                Icon = SystemIcons.Application, //TODO change to icon
                ContextMenuStrip = _trayMenu,
                Visible = true,
                Text = "Outlook2Excel"
            };

            _engine = new Outlook2Excel.Core.Engine();
            _engine.PropertyChanged += Engine_PropertyChanged;
            _engine.RunNow();

            //Hide the main window
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
        }
        private void Engine_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(_engine.LastRan))
            {
                BeginInvoke(() =>
                {
                    _lastRanItem.Text = $"Last ran: {_engine.LastRan:g}";
                });
            }
        }

        private void OnRunNow(object sender, EventArgs e) 
        {
            _engine.RunNow();
            _lastRanItem.Text = _engine.LastRan;
        }
        private void OnPause(object sender, EventArgs e)
        {
            _engine.Pause();
            _pauseItem.Visible = false;
            _unpauseItem.Visible = true;
        }
        private void OnUnPause(object sender, EventArgs e) 
        {
            _engine.UnPause();
            _pauseItem.Visible = true;
            _unpauseItem.Visible = false;
        }
        private void OnExit(object sender, EventArgs e)
        {
            _engine.Dispose();
            _trayIcon.Visible = false;
            Application.Exit();
        }
    }
}
