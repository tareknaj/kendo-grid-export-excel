using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using KendoGridExportExcel.EF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

namespace KendoGridExportExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Read([DataSourceRequest]DataSourceRequest request)
        {
            DemoEntities db = new DemoEntities();
            return Json(db.Table.ToDataSourceResult(request));
        }

        public FileResult Export([DataSourceRequest]DataSourceRequest request)
        {
            DemoEntities db = new DemoEntities();
            byte[] bytes = WriteExcel(db.Table.ToDataSourceResult(request).Data, new string[] { "Id", "Name" });

            return File(bytes,
                "application/vnd.ms-excel",
                "GridExcelExport.xls");
        }

        public byte[] WriteExcel(IEnumerable data, string[] columns)
        {
            MemoryStream output = new MemoryStream();
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IFont headerFont = workbook.CreateFont();
            headerFont.Boldweight = (short)FontBoldWeight.Bold;
            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.SetFont(headerFont);
            headerStyle.Alignment = HorizontalAlignment.Center;

            //(Optional) freeze the header row so it is not scrolled
            sheet.CreateFreezePane(0, 1, 0, 1);

            IEnumerator foo = data.GetEnumerator();
            foo.MoveNext();
            Type t = foo.Current.GetType();

            IRow header = sheet.CreateRow(0);
            PropertyInfo[] properties = t.GetProperties();
            int colIndex = 0;
            for (int i = 0; i < properties.Length; i++)
            {
                if (columns.Contains(properties[i].Name))
                {
                    ICell cell = header.CreateCell(colIndex);
                    cell.CellStyle = headerStyle;
                    cell.SetCellValue(properties[i].Name);
                    colIndex++;
                }
            }

            int rowIndex = 0;
            foreach (object o in data)
            {
                colIndex = 0;
                IRow row = sheet.CreateRow(rowIndex + 1);
                for (int i = 0; i < properties.Length; i++)
                {
                    if (columns.Contains(properties[i].Name))
                    {
                        row.CreateCell(colIndex).SetCellValue(properties[i].GetValue(o, null).ToString());
                        colIndex++;
                    }
                }
                rowIndex++;
            }

            workbook.Write(output);
            return output.ToArray();
        }
    }
}
