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

        public void SendEmail(string recipient, string cc, string subject, string body, string attachmentPath)
        {

            string currentDateTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm");

            try
            {

                QueryManager queryManager = new QueryManager();
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                Attachment attachment = new Attachment(attachmentPath);

                var emailData = GetRecipients();

                message.From = new MailAddress("vaughank@bipa.na");
                message.To.Add(recipient);
                message.Subject = subject;
                message.IsBodyHtml = false;
                message.Body = body;
                message.Attachments.Add(attachment);
                message.CC.Add(cc);
                //message.CC.Add(new MailAddress("andyc@bipa.na"));
                //message.CC.Add(new MailAddress("alueendor@bipa.na"));

                //string[] attachments;
                //foreach(string att in attachments)
                //{
                //    message.Attachments.Add(att);
                //}

                smtp.Port = 587;
                smtp.Host = "smtp.office365.com";
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential("vaughank@bipa.na", "jmnknlmskvxywvmv"); 
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");

                using (StreamWriter errorWriter = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\logs\errorLog.log", true))
                {
                    errorWriter.WriteLine($"Current Time: {currentDateTime} - Error: {ex.ToString()}");
                    errorWriter.WriteLine();
                }
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
