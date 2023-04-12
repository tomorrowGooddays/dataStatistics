using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDDataStatistics
{
    public class Constant
    {
        public static Dictionary<string, string> TitleCellDic  = new Dictionary<string, string>()
        {
            {"TBDLMJ","图斑地类面积" },
            {"ZRQDM","自然区" },
            {"PDJB","坡度" },
            {"TCHDJB","土层厚度" },
            {"TRZDJB","土壤质地" },
            {"TRYJZHLJB","土壤有机质含量" },
            {"TRPHZJB","土壤pH值" },
            {"SWDYXJB","生物多样性" },
            {"TRZJSWRJB","土壤重金属污染状况" },
            {"SZJB","熟制" },
            {"GDEJDLJB","耕地二级地类" },
            {"BHQPDJB","坡度" },
            {"BHQTCHDJB","土层厚度" },
            {"BHQTRZDJB","土壤质地" },
            {"BHQTRYJZ_1","土壤有机质含量" },
            {"BHQTRPHZJB","土壤pH值" },
            {"BHQSWDYXJB","生物多样性" },
            {"BHQTRZJS_1","土壤重金属污染状况" },
            {"BHQSZJB","熟制" },
            {"BHQGDEJDLJ","耕地二级地类" },
        };

        //所有的title
        public static List<string> TitleNameList = TitleCellDic.Keys.ToList();

        public static List<int> TitleIndexList = new List<int>();
    }
}
