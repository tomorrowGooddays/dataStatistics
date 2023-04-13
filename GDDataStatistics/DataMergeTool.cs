using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace GDDataStatistics
{
    public class DataMergeTool
    {
        private static Dictionary<string, Dictionary<string, double>> mergeDic = new Dictionary<string, Dictionary<string, double>>();
        /// <summary>
        /// dictionaryList-Merge
        /// </summary>
        /// <param name="dictList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, double>> MergeData(List<Dictionary<string, Dictionary<string, double>>> dictList)
        {
            mergeDic.Clear();

            for (int i = 0; i < dictList.Count(); i++)
            {
                MergeTwoDicData(dictList[i]);
            }

            return mergeDic;
        }

        private static void MergeTwoDicData(Dictionary<string, Dictionary<string, double>> baseDic)
        {
            for (int i = 0; i < baseDic.Count; i++)
            {
                var item = baseDic.ElementAt(i);
                string titileName = item.Key;
                Dictionary<string, double> baseItem = item.Value;

                List<string> mergeTitleNames = mergeDic.Keys.ToList();

                if (mergeTitleNames.Contains(titileName))
                {
                    //有相同的合并
                    Dictionary<string, double> mergeItem = mergeDic[titileName];

                    var dictionaries = new[] { baseItem, mergeItem };

                    mergeDic[titileName] = MergeTwoDic(dictionaries);
                }
                else
                {
                    //没有相同的就直接添加进去
                    mergeDic[titileName] = baseItem;
                }
            }
        }

        /// <summary>
        /// 合并两个字典
        /// </summary>
        /// <param name="dictionaries"></param>
        /// <returns></returns>
        public static Dictionary<string, double> MergeTwoDic(Dictionary<string, double>[] dictionaries)
        {
            return dictionaries
                              .SelectMany(d => d)
                              .GroupBy(
                                kvp => kvp.Key,
                                (key, kvps) => new { Key = key, Value = kvps.Sum(kvp => kvp.Value) }
                              )
                              .ToDictionary(x => x.Key, x => x.Value);
        }

        public static DataTable ConvertData(Dictionary<string, Dictionary<string, double>> dicData)
        {
            DataTable dt = new DataTable("myTable");

            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("Sum", typeof(double));

            for (int i = 0; i < dicData.Count; i++)
            {
                var item = dicData.ElementAt(i);

                string titleName = item.Key;
                Dictionary<string, double> itemValue = item.Value;

                // 使用LINQ按照Key排序
                var sortedDict = from entry in itemValue orderby entry.Key ascending select entry;

                for (int j = 0; j < sortedDict.ToList().Count(); j++)
                {
                    var itemJ = sortedDict.ElementAt(j);

                    if (j == 0)
                    {
                        dt.Rows.Add(titleName, itemJ.Key, itemJ.Value);
                    }
                    else
                    {
                        dt.Rows.Add("", itemJ.Key, itemJ.Value);
                    }
                }
            }

            return dt;
        }

    }
}
