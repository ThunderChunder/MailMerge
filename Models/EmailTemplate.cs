using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;

namespace MailMerge.Models
{
    public class EmailTemplate
    {
        [Required]
        public string Recipient{ get; set; }
        public string Sender{ get; set; }
        [Required]
        public string Subject { get; set; }
        [Required]
        public string Body { get; set; }
        [Required]
        public IFormFile file {get; set;}

        public EmailTemplate(){}
        public EmailTemplate(EmailTemplate emailTemplate)
        {
            this.Recipient = emailTemplate.Recipient;
            this.Subject = emailTemplate.Subject;
            this.Body = emailTemplate.Body;
        }
    }
}