using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NPOI_API.Models;

namespace NPOI_API.Controllers
{
    public class Test2Controller : Controller
    {
        // GET: Test2
        public ActionResult Index()
        {
            string path = @"C:\Users\Fawzy\Downloads\Copy of report example.xls";
            XSSFWorkbook workbook = new XSSFWorkbook(path);
            ISheet sheet = workbook.GetSheetAt(0);
            IRow Row = sheet.GetRow(12); 
            List<Table> Table = new List<Table>();
            int cellCount = Row.LastCellNum;
            for (int Raw = 0; Raw < cellCount; Raw++)
            {
                Table tb = new Table();
               // tb.id = ((IRow)Row.Ce
            }
            return View();
        }
    }
}