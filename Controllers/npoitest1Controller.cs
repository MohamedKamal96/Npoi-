using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI_API.Models;

namespace NPOI_API.Controllers
{
    public class npoitest1Controller : Controller
    {
        // GET: npoitest1
        public ActionResult Index()
        {
            string path = @"C:\Users\Fawzy\Downloads\report example (4).xls";

            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet("DIOC103");
            List<Table> tables = new List<Table>();
            for (int row = 12; row <= 26; row++)
            {
                Table tab = new Table();
                tab.SectionDescription = sheet.GetRow(row).GetCell(1).StringCellValue;
                tab.PPM = sheet.GetRow(row).GetCell(2).NumericCellValue;
                tab.MetCriteria = sheet.GetRow(row).GetCell(3).NumericCellValue;
                tables.Add(tab);
            }
            return View();
        }
    }
}