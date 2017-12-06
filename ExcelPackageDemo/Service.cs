using ExcelPackageDemo.Model;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPackageDemo
{
    public class Service
    {
        string filepath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        #region 已知的class
        public List<Park> CreateExcelData()
        {
            //讀json檔進來 已有檔案存在資料夾的情況

            string filelocate = Path.Combine(filepath + @"\File\tpepark.json");

            //檔案model
            List<Park> list = new List<Park>();
            try
            {
                StreamReader sr = new StreamReader(filelocate);
                string json = sr.ReadToEnd();
                list = JsonConvert.DeserializeObject<List<Park>>(json);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return list;

        }
        //取得要class要放在Excel的欄位名稱
        public List<string> GetExcelColumn()
        {
            List<string> column = new List<string>();
            //取得類別的ColumnNameAttribute
            var p = typeof(Park);
            var headers = p.GetProperties();
            foreach (PropertyInfo prop in headers)
            {
                //取得所有自訂屬性陣列
                object[] attrs = prop.GetCustomAttributes(true);
                foreach (var attr in attrs)
                {
                    ColumnNameAttribute customAttr = attr as ColumnNameAttribute;
                    column.Add(customAttr?.Description);
                }

            }
            return column;
        }
        public void GenerateExcelByClass()
        {
            List<Park> model = CreateExcelData();

            using (ExcelPackage ep = new ExcelPackage())
            {
                string sheetName = nameof(Park);
                ep.Workbook.Worksheets.Add(sheetName);
                ExcelWorksheet sheet = ep.Workbook.Worksheets[sheetName];

                
               

                //欄位
                List<string> column = GetExcelColumn();
                for (int i = 0; i < column.Count; i++)
                {
                    sheet.Cells[1, i + 1].Value = column[i];
                }
                //資料
                if (model.Count > 0)
                {
                    sheet.Cells["A2"].LoadFromCollection(model);
                }
                sheet.Cells.AutoFitColumns();



                FileStream fs = new FileStream(filepath + "/File/TaipeiPark.xls", FileMode.Create);
                ep.SaveAs(fs);
                fs.Close();
            }


                



        }
        #endregion

        #region 從資料庫撈資料或call api  沒有定class
        public DataTable CreateExcelData2()
        {
            //讀json檔進來 已有檔案存在資料夾的情況

            string filelocate = Path.Combine(filepath + @"\File\tpepark.json");

            //檔案model
            DataTable dt = new DataTable();
            try
            {
                StreamReader sr = new StreamReader(filelocate);
                string json = sr.ReadToEnd();
                dt = (DataTable)JsonConvert.DeserializeObject(json, typeof(DataTable));

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;

        }
        //用datatable產生Excel
        public void GenerateExcelByDataTable()
        {
            DataTable dt = CreateExcelData2();
            ExcelPackage ep = new ExcelPackage();

            ep.Workbook.Worksheets.Add("test");
            ExcelWorksheet sheet = ep.Workbook.Worksheets["test"];
            //資料
            if (dt.Rows.Count > 0)
            {
                sheet.Cells["A1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Custom);
            }
            sheet.Cells.AutoFitColumns();

            FileStream fs = new FileStream(filepath + "/File/DataTable.xls", FileMode.Create);
            ep.SaveAs(fs);
            fs.Close();
        }
        #endregion
    }
}
