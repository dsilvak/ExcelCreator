using System.Net;
using System.Net.Mail;
using System.Configuration;

namespace ExcelCreator.Services
{
    class Sendmail
    {
        public string mailAccount = ConfigurationManager.AppSettings["mailAccount"].ToString();
        public string mailPassword = ConfigurationManager.AppSettings["mailPassword"].ToString();

        public void Email(Attachment attachment, MailMessage mailMessage, string subject, string message)
        {
            var smtp = new SmtpClient("smtp.gmail.com");

            smtp.EnableSsl = true;
            smtp.Port = 587;
            smtp.Credentials = new NetworkCredential(mailAccount, mailPassword);
                        
            mailMessage.Attachments.Add(attachment);
            mailMessage.Subject = $"{subject}";
            mailMessage.From = new MailAddress(mailAccount);
            mailMessage.Body = message;

            smtp.Send(mailMessage);
        }
        public void Email(MailMessage mailMessage, string subject, string message)
        {
            var smtp = new SmtpClient("smtp.gmail.com");

            smtp.EnableSsl = true;
            smtp.Port = 587;
            smtp.Credentials = new NetworkCredential(mailAccount, mailPassword);

            mailMessage.Subject = $"{subject}";
            mailMessage.From = new MailAddress(mailAccount);
            mailMessage.Body = message;

            smtp.Send(mailMessage);
        }
    }
}