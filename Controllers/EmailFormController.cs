using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MailMerge.Models;
using MailMerge.Contracts;
using Microsoft.AspNetCore.Http;
using System.IO;
using ExcelDataReader;

namespace MailMerge.Controllers
{
    public class EmailFormController : Controller
    {
        //Service to send emails
        private readonly IMailDependency _SMTPService;
        //Service to write files
        private readonly IFileWriter _FileWriter;

        public EmailFormController(IMailDependency SMTPService, IFileWriter FileWriter)
        {
            _SMTPService = SMTPService;
            _FileWriter = FileWriter;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public async Task<IActionResult> onPost(EmailTemplate emailTemplate)
        {
            if(ModelState.IsValid)
            {
                _SMTPService.ProcessEmails(emailTemplate);

                return View("Index");
            }
            return View("Index");
        }

        [HttpPost]  
        public IActionResult uploadFile(IFormFile file)
        {
            string FileName= Guid.NewGuid().ToString() + Path.GetExtension(file.FileName);
            string SavePath = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets",FileName);

            _FileWriter.WriteFileToDisk(file, SavePath);

            return View("Index");
        }
    }
}