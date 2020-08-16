using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MailMerge.Models;

namespace MailMerge.Processors
{
    public class EmailInterpolation
    {
        public List<EmailTemplate> MailMerge(EmailTemplate emailTemplate, string[,] spreadSheet)
        {
            List<EmailTemplate> emailList = new List<EmailTemplate>();
            EmailTemplate temp = new EmailTemplate();

            for(int i = 0; i < spreadSheet.GetLength(0); i++)
            {
                for(int j = 0; j < spreadSheet.GetLength(1); j++)
                {
                    temp.Recipient = emailTemplate.Recipient.Replace("{{"+spreadSheet[0,j]+"}}", spreadSheet[i,j]);
                    temp.Subject = emailTemplate.Subject.Replace("{{"+spreadSheet[0,j]+"}}", spreadSheet[i,j]);
                    temp.Body = emailTemplate.Body.Replace("{{"+spreadSheet[0,j]+"}}", spreadSheet[i,j]);
                }
                emailList.Add(temp);
                // Regex regex = new Regex("{{"+Keys[i]+"}}");
                // temp.Recipient = regex.Replace(emailTemplate.Recipient, cell);
                // temp.Subject = regex.Replace(emailTemplate.Subject, cell);
                // temp.Body = regex.Replace(emailTemplate.Body, cell);
            }
            return emailList;
        }
    }
}