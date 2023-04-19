using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDDataStatistics
{
    public class DistrictDataInfo
    {

        public string DistrictName { get; set; }

        public string DistrictCode { get; set; }


        public Dictionary<string,Dictionary<string,double>> DistrictDataDic { get; set; }
    }
}
