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
                   // temp.Recipient = emailTemplate.Recipient.Replace("{{"+spreadSheet.Rows[0][j]+"}}", spreadSheet.Rows[i][j].ToString());
                   // temp.Subject = emailTemplate.Subject.Replace("{{"+spreadSheet.Rows[0][j]+"}}", spreadSheet.Rows[i][j].ToString());
                    temp.Body = temp.Body.Replace("{{"+spreadSheet.Rows[0][j]+"}}", spreadSheet.Rows[i][j].ToString());
                    // Regex regex = new Regex("{{"+spreadSheet.Rows[0][j].ToString()+"}}");
                    // //Console.WriteLine("{{"+spreadSheet.Rows[0][j].ToString()+"}}");
                    // //temp.Recipient = regex.Replace(eiailTemplate.Recipient, spreadSheet.Rows[i][j].ToString());
                    // //temp.Subject = regex.Replace(emailTemplate.Subject, spreadSheet.Rows[i][j].ToString());
                    // temp.Body = regex.Replace(emailTemplate.Body, spreadSheet.Rows[i][j].ToString());
                    Console.WriteLine(temp.Body);
                }
                emailList.Add(new EmailTemplate(temp));
            }
            // Regex regex = new Regex("{{"+spreadSheet.Rows[0][0].ToString()+"}}");
            // temp.Body = regex.Replace(emailTemplate.Body, spreadSheet.Rows[1][0].ToString());
            // //Regex r = new Regex("{{"+spreadSheet.Rows[0][1].ToString()+"}}");
            // temp.Subject = regex.Replace(emailTemplate.Subject, spreadSheet.Rows[1][0].ToString());
            // Console.WriteLine(temp.Subject+temp.Body+emailTemplate.Body);

            foreach(var x in emailList)
            {
                //Console.WriteLine(x.Body);
            }
            return emailList;
        }
    }
}