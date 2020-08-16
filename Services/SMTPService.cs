using System;
using System.Collections.Generic;
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
        private IExcelReader _ExcelReader; 
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
            var spreadSheet = _ExcelReader.Create2DList();
            //stores column headers from excel spreadsheet 
            var keys = spreadSheet[0].ToArray();

            for(int i = 0; i < spreadSheet.Count; i++)
            {
                for(int j = 0; j < spreadSheet.Count; j++)
                {
                    InterpolateEmail(emailTemplate, spreadSheet[j], keys);
                    SendMail(emailTemplate);
                }
            }
            
        }
        public async Task SendMail(EmailTemplate emailTemplate)
        {
            Console.WriteLine(emailTemplate.Subject + emailTemplate.Body);

            MailAddress from = new MailAddress(_Configuration["SenderEmailAddress"]);
            MailAddress to = new MailAddress(emailTemplate.Recipient);
            MailMessage message = new MailMessage(from, to);

            message.Body = emailTemplate.Body;
            message.BodyEncoding =  System.Text.Encoding.UTF8;

            message.Subject = emailTemplate.Subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;

            //await Task.Run(() => SmtpClient.SendAsync(message, ""));

            cleanUpMessage(message);
        }

        public void cleanUpMessage(MailMessage message)
        {
            message.Dispose();
        }

        public void InterpolateEmail(EmailTemplate emailTemplate, List<string> spreadSheet, string[] keys)
        {
            EmailInterpolation.MailMerge(emailTemplate, spreadSheet, keys);
        }

    }
}