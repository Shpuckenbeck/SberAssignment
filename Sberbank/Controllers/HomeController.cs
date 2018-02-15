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
        /// <summary>
        /// Домашний метод. Выводит все данные в таблицу, сортирую по дате
        /// </summary>
        /// <returns></returns>
        public IActionResult Index()
        {
            IndexViewModel model = new IndexViewModel();
            model.records = _context.Records.OrderBy(p => p.date).ToList();
            return View(model);
        }
        /// <summary>
        /// Добавление записи в таблицу с проверкой правильности введённых данных
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> addrec([FromBody] AddRecModel model)
        {
            Double parsedearn;
            float parsedcurr, parsedind;
            bool parse1, parse2, parse3;
            parse1 = Double.TryParse(model.earnings.ToString(), out parsedearn);
            parse2 = Single.TryParse(model.currency.ToString(), out parsedcurr);
            parse3 = Single.TryParse(model.index.ToString(), out parsedind);
            if (parse1 && parse2 && parse3 ) //если все колонки имеют верные данные, добавляем запись
            {
                Record rec = new Record();
                rec.date = DateTime.Now.Date;
                rec.earnings = parsedearn;
                rec.currency = parsedcurr;
                rec.index = parsedind;
                _context.Records.Add(rec);
                await _context.SaveChangesAsync();

            }
            
            IndexViewModel inmodel = new IndexViewModel();
            inmodel.records = _context.Records.ToList();
            return PartialView("table", inmodel);
        }
       /// <summary>
       /// Экспорт данных из БД в файл Excel
       /// </summary>
       /// <returns></returns>
        public FileResult Export ()
        {
            string wwwrootPath = _hostingEnvironment.WebRootPath;
            string fileName = @"Данные.xlsx";
            FileInfo file = new FileInfo(Path.Combine(wwwrootPath, fileName));

            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(wwwrootPath, fileName));
            }
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Данные");
                using (var range = worksheet.Cells[1, 1, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(00, 102, 00));
                    range.Style.Font.Color.SetColor(Color.White);
                }
                using (var range = worksheet.Cells[1, 1, 1, 2])  
                {
                    range.Merge = true;
                    range.Value = "Показатели заёмщика ";

                }
                worksheet.Cells[1, 3].Value = "Валюта";
                worksheet.Cells[1, 4].Value = "ИНДЕКСЫ";
                using (var range = worksheet.Cells[2, 1, 2, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(194, 215, 155));
                    range.Style.Font.Color.SetColor(Color.Black);
                }
                worksheet.Cells[2, 1].Value = "Дата";
                worksheet.Cells[2, 2].Value = "Выручка";
                worksheet.Cells[2, 3].Value = "серебро, руб";
                worksheet.Cells[2, 4].Value = "Индекс ММВБ Last";
                int currrow = 3;
                var data = _context.Records.OrderBy(p => p.date).ToList();
                foreach (Record rec in data)
                {
                    worksheet.Cells[currrow, 1].Value = rec.date.ToString("dd.MM.yyyy");
                    worksheet.Cells[currrow, 2].Value = rec.earnings.ToString();
                    worksheet.Cells[currrow, 3].Value = rec.currency.ToString();
                    worksheet.Cells[currrow, 4].Value = rec.index.ToString();
                    currrow++;

                }


                package.Save();

                byte[] fileBytes = System.IO.File.ReadAllBytes(wwwrootPath + @"\"+ fileName);
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
        /// <summary>
        /// Метод импорта файла Excel. После нажаитя клавиши "Импорт" выдаёт ошибку,
        /// однако операция добавления проходит, в чём можно убедиться, вручную выйдя
        /// на главную страницу. Ликвидировать ошибку не удалось
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        /// 
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
                    file.CopyTo(stream); //загрузка файла на сервер
                    //stream.Position = 0;
                
                }
                FileInfo existingfile = new FileInfo(fullPath);
                using (ExcelPackage package = new ExcelPackage(existingfile))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1]; //открытие файла для анализа
                    for (int i = sheet.Dimension.Start.Row +2; i <= sheet.Dimension.End.Row; i++)
                    {
                        Record rec = new Record();
                        DateTime parseddate;
                        Double parsedearn;
                        float parsedcurr, parsedind;
                        //Проверка на соответствие типов входных данных
                        bool parse1, parse2, parse3, parse4;
                        parse1 = DateTime.TryParse(sheet.Cells[i, 1].Value.ToString(), out parseddate); 
                        parse2 = Double.TryParse(sheet.Cells[i, 2].Value.ToString(), out parsedearn);
                        parse3 = Single.TryParse(sheet.Cells[i, 3].Value.ToString(), out parsedcurr);
                        parse4 = Single.TryParse(sheet.Cells[i, 4].Value.ToString(), out parsedind);
                        if (parse1&&parse2&&parse3&&parse4) //если все колонки имеют верные данные, добавляем запись
                        {
                            rec.date = parseddate;
                            rec.earnings = parsedearn;
                            rec.currency = parsedcurr;
                            rec.index = parsedind;
                            _context.Records.Add(rec);
                            await _context.SaveChangesAsync();

                        }                       


                    }
                }
            }
            IndexViewModel inmodel = new IndexViewModel();
            inmodel.records = _context.Records.ToList();
            return RedirectToAction("Index");
        }

        /// <summary>
        /// Генерация .csv из файла Excel
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public FileResult CSV(IFormFile file)
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
                                        

                }
                FileInfo existingfile = new FileInfo(fullPath);
                using (ExcelPackage package = new ExcelPackage(existingfile))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1]; //открытие файла для анализа      
                    var output = new StringBuilder();
                    string headers = "";
                    for (int i = 1; i < 4; i++)
                    {
                        headers += sheet.Cells[2, i].Value.ToString() + ", ";
                    }
                    headers += sheet.Cells[2, 4].Value.ToString();
                    output.AppendLine(headers);
                    for (int i = sheet.Dimension.Start.Row + 2; i <= sheet.Dimension.End.Row; i++)
                    {                       
                        DateTime parseddate = DateTime.Now;
                        Double parsedearn;
                        float parsedcurr, parsedind;
                        bool parse1, parse2, parse3, parse4;
                        //parse1 = DateTime.TryParse(sheet.Cells[i, 1].Value.ToString(), out parseddate);
                        parse2 = Double.TryParse(sheet.Cells[i, 2].Value.ToString(), out parsedearn);
                        parse3 = Single.TryParse(sheet.Cells[i, 3].Value.ToString(), out parsedcurr);
                        parse4 = Single.TryParse(sheet.Cells[i, 4].Value.ToString(), out parsedind);
                        if (parse2 && parse3 && parse4) 
                        {
                            output.AppendLine(DateTime.Now.Date.ToString() + ", " + parsedearn.ToString() + ", " + parsedcurr.ToString() + ", " + parsedind.ToString());
                        }
                        


                    }
                    byte[] buffer = Encoding.GetEncoding(1251).GetBytes(output.ToString());
                    return File(buffer, "text/csv", $"Данные.csv");
                }
            }

            return null;

        }
        
    }
}
