using System;
using System.Net.Mail;
using MailMerge.Contracts;
using MailMerge.Models;
using MailMerge.Processors;
using Microsoft.Extensions.Configuration;

namespace MailMerge.Services
{
    public class SMTPService : IMailDependency
    {
        //Service to get appsettings.json configuration values
        private readonly IConfiguration _Configuration;
        private SmtpClient SmtpClient;
        private EmailInterpolation EmailInterpolation;
        public SMTPService(IConfiguration Configuration)
        {
            _Configuration = Configuration;
            //Sets Smtp Server to outlook domain as specified in appsettings
            SmtpClient = new SmtpClient(_Configuration["OutLookSMTPServer"]);
            EmailInterpolation = new EmailInterpolation();
        }
        public void ProcessEmails(EmailTemplate emailTemplate)
        {
            InterpolateEmail(emailTemplate);
            SendMail(emailTemplate);
        }
        public void SendMail(EmailTemplate emailTemplate)
        {
            InterpolateEmail(emailTemplate);

            Console.WriteLine(emailTemplate.Subject + emailTemplate.Body);

            MailAddress from = new MailAddress(_Configuration["SenderEmailAddress"]);
            MailAddress to = new MailAddress(emailTemplate.Recipient);
            MailMessage message = new MailMessage(from, to);

            message.Body = emailTemplate.Body;
            message.BodyEncoding =  System.Text.Encoding.UTF8;

            message.Subject = emailTemplate.Subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;

            SmtpClient.SendAsync(message, "");

            cleanUpMessage(message);
        }

        public void cleanUpMessage(MailMessage message)
        {
            message.Dispose();
        }

        public void InterpolateEmail(EmailTemplate emailTemplate)
        {
            EmailInterpolation.MailMerge(emailTemplate);
        }

    }
}