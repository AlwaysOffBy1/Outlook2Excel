using Outlook2Excel.Core;

namespace Outlook2Excel.GUI
{
    public partial class Form1 : Form
    {
        private NotifyIcon _trayIcon;
        private ContextMenuStrip _trayMenu;
        private Outlook2Excel.Core.Engine _engine;
        public Form1()
        {
            InitializeComponent();

            _trayMenu = new ContextMenuStrip();
            _trayMenu.Items.Add("Run Now", null, OnRunNow);
            _trayMenu.Items.Add("Pause", null, OnPause);
            _trayMenu.Items.Add("Exit", null, OnExit);

            _trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application, //TODO change to icon
                ContextMenuStrip = _trayMenu,
                Visible = true,
                Text = "Outlook2Excel"
            };

            _engine = new Outlook2Excel.Core.Engine();

            //Hide the main window
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
        }

        private void OnRunNow(object sender, EventArgs e) => _engine.RunNow();
        private void OnPause(object sender, EventArgs e) => _engine.Pause();
        private void OnExit(object sender, EventArgs e)
        {
            _engine.Dispose();
            _trayIcon.Visible = false;
            Application.Exit();
        }
    }
}
