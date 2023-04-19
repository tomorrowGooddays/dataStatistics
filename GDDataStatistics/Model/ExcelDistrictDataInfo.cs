using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDDataStatistics.Model
{
    public class ExcelDistrictDataInfo
    {
        /// <summary>
        /// excel的fileName
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 按照行政区域统计的结果
        /// </summary>
        public List<DistrictDataInfo> DistrictDataInfos { get; set; } = new List<DistrictDataInfo>();

    }
}
