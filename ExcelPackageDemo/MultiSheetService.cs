using ExcelPackageDemo.Model;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPackageDemo
{
    public class MultiSheetService
    {
        static string filepath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        #region 已知的class

        public static void GenerateMultiSheetExcel()
        {


            ExcelPackage ep = new ExcelPackage();
             ep.Workbook.Worksheets.Add("Sheet1");
            ep.Workbook.Worksheets.Add("Sheet2");//第二個分頁

            //取得剛加入的實體
            ExcelWorksheet sheet = ep.Workbook.Worksheets["Sheet1"];

            //標題
            sheet.Cells[1, 1].Value = "姓名";
            sheet.Cells[1, 2].Value = "年齡";
            sheet.Cells[1, 3].Value = "身高";
            //資料列
            sheet.Cells["A2"].Value = "John";
            sheet.Cells["B2"].Value = "20";
            sheet.Cells["C2"].Value = "180";
            sheet.Cells["A3"].Value = "Emily";
            sheet.Cells["B3"].Value = "25";
            sheet.Cells["C3"].Value = "160";

            //合併儲存格
            sheet.Cells[4, 1, 4, 3].Merge = true;
            sheet.Cells[4, 1].Value = "This is merge row";

            //註解
            sheet.Cells["A2"].AddComment("employee's Name", "Kuan");


            //Format the header
            using (ExcelRange rng = sheet.Cells["A1:C1"])
            {
                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//文字置中
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;          //設定背景實線            
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                rng.Style.Font.Color.SetColor(Color.White);
                rng.Style.Font.SetFromFont(new Font("Consolas", 10, FontStyle.Italic | FontStyle.Bold));//粗體 斜體
            }
            //自動調寬
            sheet.Cells.AutoFitColumns();

            //存檔
            FileStream fs = new FileStream(filepath + "/File/MultiSheet.xls", FileMode.Create);
            ep.SaveAs(fs);
            fs.Close();







        }
        #endregion
    }
}
