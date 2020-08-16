using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using MailMerge.Models;
using MailMerge.Contracts;
using Microsoft.AspNetCore.Http;
using System.IO;

namespace MailMerge.Controllers
{
    public class EmailFormController : Controller
    {
        //Service to send emails
        private readonly IMailDependency _SMTPService;
        //Service to write files
        private readonly IFileWriter _FileWriter;

        public EmailFormController(IMailDependency SMTPService, IFileWriter fileWriter, IExcelReader excelReader)
        {
            _SMTPService = SMTPService;
            _FileWriter = fileWriter;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult onPost(EmailTemplate emailTemplate)
        {
            if(ModelState.IsValid)
            {
                //string FileName= Guid.NewGuid().ToString() + Path.GetExtension(file.FileName);
                string SavePath = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets","spreadsheet.xlsx");

                _FileWriter.WriteFileToDisk(emailTemplate.file, SavePath);
                _SMTPService.ProcessEmails(emailTemplate);

                return View("Index");
            }
            return View("Index");
        }
    }
}