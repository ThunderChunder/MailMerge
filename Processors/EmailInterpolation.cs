using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MailMerge.Models;

namespace MailMerge.Processors
{
    public class EmailInterpolation
    {
        public EmailTemplate MailMerge(EmailTemplate emailTemplate, List<string> spreadSheet, string[] Keys)
        {
            EmailTemplate temp = new EmailTemplate();
            int i = 0;
            foreach(var cell in spreadSheet)
            {
                // Regex regex = new Regex("{{"+Keys[i]+"}}");
                // temp.Recipient = regex.Replace(emailTemplate.Recipient, cell);
                // temp.Subject = regex.Replace(emailTemplate.Subject, cell);
                // temp.Body = regex.Replace(emailTemplate.Body, cell);
                temp.Recipient = emailTemplate.Recipient.Replace("{{"+Keys[i]+"}}", cell);
                temp.Subject = emailTemplate.Subject.Replace("{{"+Keys[i]+"}}", cell);
                temp.Body = emailTemplate.Body.Replace("{{"+Keys[i]+"}}", cell);
                i++;
            }
            return temp;
        }
    }
}