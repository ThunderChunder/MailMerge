using System.ComponentModel.DataAnnotations;

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
    }
}