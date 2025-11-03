using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService
{
    public class EmailService
    {
        public async Task SendEmail(string mailTo, Email objEmail, List<string> listBCC, string excelFilePath = null)
        {
            try
            {
                string host = ConfigurationManager.AppSettings["SmtpHost"];
                string userName = ConfigurationManager.AppSettings["SmtpUserName"];
                string password = ConfigurationManager.AppSettings["SmtpPassword"];
                int port = Convert.ToInt32(ConfigurationManager.AppSettings["SmtpPort"]);

                SmtpClient client = new SmtpClient(host)
                {
                    UseDefaultCredentials = true,
                    Credentials = new NetworkCredential(userName, password),
                    Port = port,
                    EnableSsl = true
                };

                MailMessage msg = new MailMessage
                 {
                    From = new MailAddress(userName),
                    Subject = objEmail.subject,
                    Body = objEmail.body,
                    IsBodyHtml = true,
                    Priority = MailPriority.High
                };

                msg.To.Add(mailTo);

                foreach (var bcc in listBCC)
                {
                    msg.Bcc.Add(bcc);
                }

                if (!string.IsNullOrEmpty(excelFilePath) && File.Exists(excelFilePath))
                {
                    Attachment attachment = new Attachment(excelFilePath);
                    msg.Attachments.Add(attachment);
                }

                 client.Send(msg);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
