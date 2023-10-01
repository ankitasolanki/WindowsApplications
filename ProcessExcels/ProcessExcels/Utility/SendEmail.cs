using System;
using System.Net;
using System.Net.Mail;

namespace ProcessExcels.Utility
{
    public class SendEmail : ISendEmail
    {
        private string smtpserver;
        private int smtpport;
        private string smtpusername;
        private string smtppassword;
        private string senderemail;
        private string recipientemail;
        private string subject;
        private string body;
        private string attachmentpath;

        public string SmtpServer { get => this.smtpserver; set => this.smtpserver = value; }
        public int SmtpPort { get => this.smtpport; set => this.smtpport = value; }
        public string SmtpUserName { get => this.smtpusername; set => this.smtpusername = value; }
        public string SmtpPassword { get => this.smtppassword; set => this.smtppassword = value; }
        public string SenderEmail { get => this.senderemail; set => this.senderemail = value; }
        public string RecipientEmail { get => this.recipientemail; set => this.recipientemail = value; }
        public string Subject { get => this.subject; set => this.subject = value; }
        public string Body { get => this.body; set => this.body = value; }
        public string AttachmentPath { get => this.attachmentpath; set => this.attachmentpath = value; }

        public void Send()
        {
            try
            {
                using (SmtpClient smtpClient = new SmtpClient(smtpserver))
                {
                    smtpClient.Port = this.smtpport;
                    smtpClient.Credentials = new NetworkCredential(this.smtpusername, this.smtppassword);
                    smtpClient.EnableSsl = true;

                    using (MailMessage mailMessage = new MailMessage(this.senderemail, this.recipientemail))
                    {
                        mailMessage.Subject = this.subject;
                        mailMessage.Body = this.body;

                        Attachment attachment = new Attachment(this.attachmentpath);
                        mailMessage.Attachments.Add(attachment);
                        smtpClient.Send(mailMessage);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
