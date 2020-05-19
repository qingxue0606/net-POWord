using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;

namespace NetCoreSamples5.Controllers
{
    public class PageOfficeController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public PageOfficeController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        [Route("/PageOffice/POserver")]
        [Route("/PageOffice/pageoffice.js")]
        [Route("/PageOffice/pobstyle.css")]
        [Route("/PageOffice/posetup.exe")]
        [Route("/PageOffice/sealsetup.exe")]
        public ActionResult POServer()
        {
            PageOfficeNetCore.POServer.Server poServer = new PageOfficeNetCore.POServer.Server(Request, Response);
            poServer.LicenseFilePath = _webHostEnvironment.ContentRootPath + "/Views/PageOffice/";
            poServer.Run();
            return Content("OK");
        }

    }
}