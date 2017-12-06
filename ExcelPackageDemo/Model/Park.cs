using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPackageDemo.Model
{
   public class Park
    {
        [ColumnName("景點所在公園名稱")]
        public string ParkName { get; set; }
        [ColumnName("景點名稱")]
        public string Name { get; set; }
        [ColumnName("景點建造年份")]
        public string YearBuilt { get; set; }
        [ColumnName("景點開放時間")]
        public string OpenTime { get; set; }
        [ColumnName("景點圖片")]
        public string Image { get; set; }
        [ColumnName("景點介紹說明")]
        public string Introduction { get; set; }
    }
}
