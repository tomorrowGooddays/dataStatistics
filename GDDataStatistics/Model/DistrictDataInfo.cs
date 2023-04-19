using System.Collections.Generic;

namespace GDDataStatistics.Model
{
    public class DistrictDataInfo
    {
        public string DistrictCode { get; set; }

        public string DistrictName { get; set; }

        public Dictionary<string, Dictionary<string, double>> CellValueDic { get; set; }

    }
}
