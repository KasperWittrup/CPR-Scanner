using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDPR_Scanner
{
    class MailHandler
    {
        public bool sendMailToEmail(string email, string header, string content)
        {
            //E:\Websites\Sitecore\Website
            //string mailBody = File.ReadAllText("E:\\Websites\\Sitecore\\Website\\maillayout\\mail.html");
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.To.Add(email);
            message.Subject = header;
            message.From = new System.Net.Mail.MailAddress("no-reply@mariagerfjord.dk");
            message.Body = content;
            message.IsBodyHtml = true;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("192.168.207.165");
            smtp.Send(message);

            return true;
        }
    }
}
