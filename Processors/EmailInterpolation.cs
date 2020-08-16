using System.Text.RegularExpressions;
using MailMerge.Models;

namespace MailMerge.Processors
{
    public class EmailInterpolation
    {
        public EmailTemplate MailMerge(EmailTemplate emailTemplate)
        {
            string hello = "{{hello}}";

            Regex regex = new Regex(hello);
            emailTemplate.Subject = regex.Replace(emailTemplate.Body, "this");
            emailTemplate.Body = regex.Replace(emailTemplate.Body, "this");

            return emailTemplate;
        }
    }
}