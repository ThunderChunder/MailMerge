using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using MailMerge.Models;

namespace MailMerge.Processors
{
    public class EmailInterpolation
    {
        public List<EmailTemplate> MailMerge(EmailTemplate emailTemplate, DataTable spreadSheet)
        {
            List<EmailTemplate> emailList = new List<EmailTemplate>();
            EmailTemplate temp = new EmailTemplate(emailTemplate);

             for (var i = 1; i < spreadSheet.Rows.Count; i++)
            {
                for (var j = 0; j < spreadSheet.Columns.Count; j++)
                {
                    temp.Recipient = temp.Recipient.Replace("{{"+spreadSheet.Rows[0][j]+"}}", spreadSheet.Rows[i][j].ToString());
                    temp.Subject = temp.Subject.Replace("{{"+spreadSheet.Rows[0][j]+"}}", spreadSheet.Rows[i][j].ToString());
                    temp.Body = temp.Body.Replace("{{"+spreadSheet.Rows[0][j]+"}}", spreadSheet.Rows[i][j].ToString());
                }
                emailList.Add(new EmailTemplate(temp));
                temp = new EmailTemplate(emailTemplate);
            }
            return emailList;
        }
    }
}