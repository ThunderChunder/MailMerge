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
                var path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets", "f1df800b-1eef-4c8a-9a2d-fad734537178.xlsx");
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read()) //Each ROW
                            {
                                for (int column = 0; column < reader.FieldCount; column++)
                                {
                                    //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
                                    Console.WriteLine(reader.GetValue(column));//Get Value returns object
                                }
                            }
                        } while (reader.NextResult()); //Move to NEXT SHEET

                    }
                }
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