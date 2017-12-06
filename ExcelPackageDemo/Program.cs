using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPackageDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Service ss = new Service();
            Console.WriteLine("已知資料類型");
            ss.GenerateExcelByClass();
            Console.WriteLine("未知資料或從資料庫取得的");
            ss.GenerateExcelByDataTable();
            Console.WriteLine("多個分頁 & 各種用法");
            MultiSheetService.GenerateMultiSheetExcel();
            


        }
    }
}
