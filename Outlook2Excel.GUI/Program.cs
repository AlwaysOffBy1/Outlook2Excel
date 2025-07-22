using log4net;
using log4net.Config;
using log4net.Repository;
using Microsoft.VisualBasic.Logging;

namespace Outlook2Excel.GUI
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Directory.CreateDirectory("Logs");
            log4net.Config.XmlConfigurator.Configure(new FileInfo("log4net.config"));
            Application.Run(new Form1());
        }
    }
}