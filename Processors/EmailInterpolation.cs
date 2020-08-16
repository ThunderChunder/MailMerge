using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MailMerge.Models;

namespace MailMerge.Processors
{
    public class EmailInterpolation
    {
        public void MailMerge(EmailTemplate emailTemplate, List<List<string>> spreadSheet)
        {
            //column keys of spreadsheet
            string[] Keys = spreadSheet[0].ToArray();

            for(int i = 1; i < spreadSheet[0].Count; i++)
            {
                for(int j = 0; j < spreadSheet[1].Count; j++)
                {
                    // Console.WriteLine(Keys[i]);
                    Regex regex = new Regex("{{"+Keys[j]+"}}");
                    emailTemplate.Recipient = regex.Replace(emailTemplate.Recipient, spreadSheet[i][j]);
                    emailTemplate.Subject = regex.Replace(emailTemplate.Subject, spreadSheet[i][j]);
                    emailTemplate.Body = regex.Replace(emailTemplate.Body, spreadSheet[i][j]);
                }
                
            }
        }
    }
}