using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MailMerge.Models;
using MailMerge.Contracts;

namespace MailMerge.Controllers
{
    public class EmailFormController : Controller
    {
        //Service to send emails
        private readonly IMailDependency _SMTPService;

        public EmailFormController(IMailDependency SMTPService)
        {
            _SMTPService = SMTPService;
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
    }
}