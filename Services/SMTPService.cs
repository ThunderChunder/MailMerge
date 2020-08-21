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
                try
                {
                    MailAddress from = new MailAddress(_Configuration["SenderEmailAddress"]);
                    MailAddress to = new MailAddress(email.Recipient);
                    var message = new MailMessage(from, to)
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