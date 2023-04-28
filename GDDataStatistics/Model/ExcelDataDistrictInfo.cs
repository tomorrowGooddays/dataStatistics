using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDDataStatistics.Model
{
    public class ExcelDataDistrictInfo
    {
        public string FileName { get; set; }

        public Dictionary<string, Dictionary<string, double>> DataDic{ get; set; } = new Dictionary<string, Dictionary<string, double>>();

        public Dictionary<string, Dictionary<string, Dictionary<string, double>>> DataDistrictList { get; set; } = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();

        public Dictionary<string, string> DistrictNameAndCodeMap { get; set; }

        public Dictionary<string, double> DistrictTotalAmount { get; set; }
    }
}
