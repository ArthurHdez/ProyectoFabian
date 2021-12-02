using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportEXCEL
{
    public class ExcelViewModel
    {
        public string Part_Number { get; set; }
        public string Supplier_part_number { get; set; }
        public string Description_Material { get; set; }
        public int Supplier { get; set; }
        public int Min_Stock { get; set; }
        public int Max_Stock { get; set; }
    }
}
