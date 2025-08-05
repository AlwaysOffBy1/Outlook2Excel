using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Outlook2Excel.Core
{
    public static class StaticMethods
    {
        private static bool _quitting = false;
        public static void Quit(string reason, int errorCode, Exception? e)
        {
            Outlook2Excel.Core.AppLogger.Log.Error(reason, e);
            if (!_quitting && AppSettings.IsOnErrorSendEmail) SendEmail(AppSettings.OnErroSendEmailTo,AppSettings.OnErrorSendEmailSubject, reason);
            Console.WriteLine(reason);
            Environment.Exit(errorCode);
        }
        public static void SendEmail(string[] emails,  string subject, string body)
        {
            _quitting = true;
            using (var message = new MailMessage())
            {

                message.From = new MailAddress(AppSettings.OnErrorSendEmailFrom, AppSettings.OnErrorSendEmailFromName);

                message.Subject = subject;
                try
                {
                    foreach (var emailAddress in emails)
                        message.To.Add(emailAddress.Trim());
                }
                catch
                {
                    Quit("INVALID EMAIL", 600, null);
                }
                message.Body = $"<pre style=\"font-family:Lucida Console\">{body}</pre>"; //make monospace
                message.IsBodyHtml = true;

                try
                {
                    using (var smtp = new SmtpClient(AppSettings.OnErrorSendEmailSMTPPath))
                    {
                        smtp.UseDefaultCredentials = false;
                        smtp.Port = 25;
                        smtp.EnableSsl = false;


                        smtp.Send(message);
                    }
                }
                catch (Exception ex)
                {
                    AppLogger.Log.Error("Could not send email", ex);
                }

            }
        }
    }
}
