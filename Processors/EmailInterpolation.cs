using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MailMerge.Models;

namespace MailMerge.Processors
{
    public class EmailInterpolation
    {
        public void MailMerge(EmailTemplate emailTemplate, List<string> spreadSheet, string[] Keys)
        {
            var i = 0;
            foreach(var cell in spreadSheet)
            {
                // Console.WriteLine(Keys[i]);
                Regex regex = new Regex("{{"+Keys[i]+"}}");
                emailTemplate.Recipient = regex.Replace(emailTemplate.Recipient, spreadSheet[i]);
                emailTemplate.Subject = regex.Replace(emailTemplate.Subject, spreadSheet[i]);
                emailTemplate.Body = regex.Replace(emailTemplate.Body, spreadSheet[i]);
                i++;
            }
        }
    }
}