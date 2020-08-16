using System.IO;
using MailMerge.Contracts;
using Microsoft.AspNetCore.Http;

namespace MailMerge.Services
{
    public class FileWriterService : IFileWriter
    {
        public void WriteFileToDisk(IFormFile file, string SavePath)
        {
            if (file!=null)
            {
                using(var stream = new FileStream(SavePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
            }
        }
    }
}