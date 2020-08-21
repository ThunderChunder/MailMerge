using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Net.Mail;
using System.Threading.Tasks;
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
        private readonly IExcelReader _ExcelReader; 
        public SMTPService(IConfiguration Configuration, IExcelReader excelReader)
        {
            _Configuration = Configuration;
            //Sets Smtp Server to outlook domain as specified in appsettings
            SmtpClient = GetMailClient();
            //Used in mail merge, replace {{ var }} with excel values
            EmailInterpolation = new EmailInterpolation();
            //Used to read excel spreadsheet and return 2D list
            _ExcelReader = excelReader;
        }
        public SmtpClient GetMailClient()
        {
            SmtpClient client = new SmtpClient(_Configuration["OutLookSMTPServer"]);
            client.Port = 587;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(_Configuration["SenderEmailAddress"], _Configuration["SenderEmailPassword"]);
            return client;
        }
        public void ProcessEmails(EmailTemplate emailTemplate)
        {
            var dataSet = _ExcelReader.CreateDataSet();
            var spreadSheet = dataSet.Tables[0];//Tables is array of sheets from excel file
           
            var interpolatedEmailList = InterpolateEmail(emailTemplate, spreadSheet);

            SendMail(interpolatedEmailList);
        }
        public void SendMail(List<EmailTemplate> emailTemplateList)
        {
            foreach(var email in emailTemplateList)
            {
                MailAddress from = CreateMailAddress(_Configuration["SenderEmailAddress"]);
                MailAddress to = CreateMailAddress(email.Recipient);
                MailMessage message = CreateMailMessage(from, to);

                SetMessageSubject(message, email.Subject);
                SetSubjectEncoding(message, System.Text.Encoding.UTF8);
                SetMessageBody(message, email.Body);
                SetBodyEncoding(message, System.Text.Encoding.UTF8);

                SmtpClient.Send(message);

                CleanUpMessage(message);
                break;
            }
        }
        public void SetMessageSubject(MailMessage message, string subject)
        {
            message.Subject = subject;
        }

        public void SetSubjectEncoding(MailMessage message, System.Text.Encoding encoding)
        {
            message.SubjectEncoding =  encoding;
        }

        public void SetBodyEncoding(MailMessage message, System.Text.Encoding encoding)
        {
            message.BodyEncoding =  encoding;
        }
        public void SetMessageBody(MailMessage message, string body)
        {
            message.Body = body;
        }

        public MailAddress CreateMailAddress(string address)
        {
            return new MailAddress(address);
        }

        public MailMessage CreateMailMessage(MailAddress from, MailAddress to)
        {
            return new MailMessage(from, to);
        }

        public void CleanUpMessage(MailMessage message)
        {
            message.Dispose();
        }

        public List<EmailTemplate> InterpolateEmail(EmailTemplate emailTemplate, DataTable spreadSheet)
        {
            return EmailInterpolation.MailMerge(emailTemplate, spreadSheet);
        }

    }
}