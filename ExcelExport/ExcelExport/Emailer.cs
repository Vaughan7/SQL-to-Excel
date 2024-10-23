using DocumentFormat.OpenXml.Wordprocessing;
using log4net.Repository.Hierarchy;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ExcelExport
{
    internal class Emailer
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void SendEmail(string recipient, string subject, string body, string attachmentPath)
        {
            try
            {
                QueryManager queryManager = new QueryManager();
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                Attachment attachment = new Attachment(attachmentPath);

                var emailData = GetRecipients();

                message.From = new MailAddress("vaughank@bipa.na");
                message.To.Add(new MailAddress(recipient));
                message.Subject = subject;
                message.IsBodyHtml = false;
                message.Body = body;
                message.Attachments.Add(attachment);

                //string[] attachments;
                //foreach(string att in attachments)
                //{
                //    message.Attachments.Add(att);
                //}

                smtp.Port = 587;
                smtp.Host = "smtp.office365.com";
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential("vaughank@bipa.na", "Bipa4321");
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
            }
            catch
            {

            }
        }

        //public Dictionary<string, object> GetRecipients()
        public EmailValueSet GetRecipients()
        {
            string jsonString = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "../../../../config/ReportInfo.json");

            var valueSet = JsonConvert.DeserializeObject<EmailWrapper>(jsonString).emailValueSet;

            return valueSet;

        }
    }
}
