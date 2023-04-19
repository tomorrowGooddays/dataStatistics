using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDDataStatistics
{
    public class ExcelDataInfo
    {
        public string FileName { get; set; }

        public Dictionary<string, Dictionary<string, double>> DataList { get; set; } = new Dictionary<string, Dictionary<string, double>>();
    }
}
