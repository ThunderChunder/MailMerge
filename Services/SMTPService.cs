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
            return new SmtpClient(_Configuration.GetValue<string>("MailSettings:OutLookSMTPServer"))
            {
                Port = _Configuration.GetValue<int>("MailSettings:Port"),
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                EnableSsl = _Configuration.GetValue<bool>("MailSettings:EnableSSL"),
                Credentials = new System.Net.NetworkCredential(
                    _Configuration.GetValue<string>("MailSettings:SenderEmailAddress"), 
                    _Configuration.GetValue<string>("MailSettings:SenderEmailPassword")
                )
            };
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
                try
                {
                    MailAddress from = new MailAddress(
                        _Configuration.GetValue<string>("MailSettings:SenderEmailAddress")
                    );
                    MailAddress to = new MailAddress(email.Recipient);
                    MailMessage message = new MailMessage(from, to)
                    {
                        Subject = email.Subject,
                        SubjectEncoding =  System.Text.Encoding.UTF8,
                        Body = email.Body,
                        BodyEncoding = System.Text.Encoding.UTF8
                    };

                    SmtpClient.Send(message);

                    message.Dispose();
                }
                catch (Exception e){Console.WriteLine(e.Message);}
            }
        }
        public List<EmailTemplate> InterpolateEmail(EmailTemplate emailTemplate, DataTable spreadSheet)
        {
            return EmailInterpolation.MailMerge(emailTemplate, spreadSheet);
        }

    }
}