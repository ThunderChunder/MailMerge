using System;
using System.Collections.Generic;
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
            SmtpClient = new SmtpClient(_Configuration["OutLookSMTPServer"]);
            //Used in mail merge, replace {{ var }} with excel values
            EmailInterpolation = new EmailInterpolation();
            //Used to read excel spreadsheet and return 2D list
            _ExcelReader = excelReader;
        }
        public void ProcessEmails(EmailTemplate emailTemplate)
        {
            var dataSet = _ExcelReader.CreateDataSet();
            var spreadSheet = dataSet.Tables[0];//Tables is array of sheets from excel file
           
            var interpolatedEmailList = InterpolateEmail(emailTemplate, spreadSheet);

            SendMail(interpolatedEmailList);
        }
        public async Task SendMail(List<EmailTemplate> emailTemplateList)
        {
            foreach(var email in emailTemplateList)
            {
                //Console.WriteLine(email.Recipient +email.Subject + email.Body);

                MailAddress from = CreateMailAddress(_Configuration["SenderEmailAddress"]);
                MailAddress to = CreateMailAddress(email.Recipient);
                MailMessage message = CreateMailMessage(from, to);

               
                SetMessageSubject(message, email.Subject);
                SetSubjectEncoding(message, System.Text.Encoding.UTF8);
                SetMessageBody(message, email.Body);
                SetBodyEncoding(message, System.Text.Encoding.UTF8);

                Console.WriteLine("To: {0} \nSubject: {1} \nBody: {2}\n", message.To, message.Subject, message.Body);

                //await Task.Run(() => SmtpClient.SendAsync(message, ""));

                CleanUpMessage(message);
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