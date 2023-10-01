using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessExcels.Utility
{
    public interface ISendEmail
    {
        string SmtpServer { get; set; }
        int SmtpPort { get; set; }
        string SmtpUserName { get; set; }
        string SmtpPassword { get; set; }
        string SenderEmail { get; set; }
        string RecipientEmail { get; set; }
        string Subject { get; set; }
        string Body { get; set; }
        string AttachmentPath { get; set; }

        void Send();
    }
}
