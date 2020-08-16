using MailMerge.Models;

namespace MailMerge.Contracts
{
    public interface IMailDependency
    {
         public void ProcessEmails(EmailTemplate EmailTemplate);
    }
}