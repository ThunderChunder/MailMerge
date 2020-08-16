using Microsoft.AspNetCore.Http;

namespace MailMerge.Contracts
{
    public interface IFileWriter
    {
        public void WriteFileToDisk(IFormFile file, string SavePath);
    }
}