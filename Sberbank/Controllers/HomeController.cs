using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Sberbank.Models;
using Sberbank.Data;
using Microsoft.AspNetCore.Http;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using Microsoft.AspNetCore.Hosting;
namespace Sberbank.Controllers
{
    public class HomeController : Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;
        private readonly BankContext _context;
        public HomeController(BankContext context, IHostingEnvironment hostingEnvironment)
        {
            _context = context;
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult Index()
        {
            IndexViewModel model = new IndexViewModel();
            model.records = _context.Records.ToList();
            return View(model);
        }
        [HttpPost]
        public async Task<IActionResult> addrec([FromBody] AddRecModel model)
        {
            Record rec = new Record();
            rec.date = DateTime.Now.Date;
            rec.currency = Convert.ToSingle(model.currency);
            rec.earnings = Convert.ToDouble(model.earnings);
            rec.index = Convert.ToSingle(model.index);
            if (TryValidateModel(rec))
            {
                _context.Records.Add(rec);
               await _context.SaveChangesAsync();

            }
            IndexViewModel inmodel = new IndexViewModel();
            inmodel.records = _context.Records.ToList();
            return PartialView("table", inmodel);
        }

        [HttpPost]
        public async Task<IActionResult> Import(IFormFile file)
        {
            string folderName = "Upload";
            string webRootPath = _hostingEnvironment.WebRootPath;
            string newPath = Path.Combine(webRootPath, folderName);
            if (!Directory.Exists(newPath))
            {
                Directory.CreateDirectory(newPath);
            }
            if (file.Length > 0)
            {
                string sFileExtension = Path.GetExtension(file.FileName).ToLower();
                string fullPath = Path.Combine(newPath, file.FileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                    //stream.Position = 0;
                
                }
                FileInfo existingfile = new FileInfo(fullPath);
                using (ExcelPackage package = new ExcelPackage(existingfile))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                    for (int i = sheet.Dimension.Start.Row + 2; i <= sheet.Dimension.End.Row; i++)
                    {
                        Record rec = new Record();
                        //rec.date = Convert.ToDateTime(sheet.Cells[i, 1].Value);
                        rec.date = DateTime.Now.Date;
                        rec.earnings = Convert.ToDouble(sheet.Cells[i, 2].Value);
                        rec.currency = Convert.ToSingle(sheet.Cells[i, 3].Value);
                        rec.index = Convert.ToSingle(sheet.Cells[i, 4].Value);

                        _context.Records.Add(rec);
                        await _context.SaveChangesAsync();


                    }
                }
            }
            IndexViewModel inmodel = new IndexViewModel();
            inmodel.records = _context.Records.ToList();
            return PartialView("table", inmodel);
        }

        
    }
}
