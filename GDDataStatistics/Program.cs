using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GDDataStatistics
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //List<Dictionary<string, Dictionary<string, double>>> dictList = new List<Dictionary<string, Dictionary<string, double>>>()
            //{
            //    new Dictionary<string, Dictionary<string, double>>()
            //    {
            //        { "a", new Dictionary<string, double>() { {"a1", 2 } } },
            //        { "b", new Dictionary<string, double>() { {"b1", 2.56 } } },
            //        { "c", new Dictionary<string, double>() { {"c1", 10.25 } } },
            //    },
            //    new Dictionary<string, Dictionary<string, double>>()
            //    {
            //        { "a", new Dictionary<string, double>() { {"a1", 6 } } },
            //        { "b", new Dictionary<string, double>() { {"b1", 3 },{"b2",40 } } },
            //        { "d", new Dictionary<string, double>() { {"d1", 100 } } },
            //    },

            //};

            //var result = DataMergeTool.MergeData(dictList);

            //DataTable dt = DataMergeTool.ConvertData(result);


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
