using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPackageDemo.Model
{
    //給Excel用的欄位名稱
   public class ColumnNameAttribute:Attribute
    {
        public string Description { get; set; }
        public ColumnNameAttribute(string desc)
        {
            this.Description = desc;
        }
        public override string ToString()
        {
            return this.Description.ToString();
        }
    }
}
