using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
            DataTable dt = new DataTable("GD");

            dt.Columns.Add("分类指标", typeof(string));
            dt.Columns.Add("指标分级", typeof(string));
            dt.Columns.Add("合计", typeof(double));

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


        public static void ExportData(string filePathAndName, Dictionary<string, Dictionary<string, double>> dicData)
        {
            if (string.IsNullOrEmpty(filePathAndName)) return;

            filePathAndName = filePathAndName.Replace(".xlsx", $"{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx");

            //创建workbook，说白了就是在内存中创建一个Excel文件
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheetGD = workbook.CreateSheet("GD");
            ISheet sheetHF = workbook.CreateSheet("HF");

            IRow rowGD = sheetGD.CreateRow(0);//添加第1行,注意行列的索引都是从0开始的

            ICell cell1GD = rowGD.CreateCell(0);//给第1行添加第1个单元格
            cell1GD.SetCellValue("分类指标");
            ICell cell2GD = rowGD.CreateCell(1);//给第1行添加第1个单元格
            cell2GD.SetCellValue("指标分级");
            ICell cell3GD = rowGD.CreateCell(2);//给第1行添加第1个单元格
            cell3GD.SetCellValue("四川盆地");
            ICell cell4GD = rowGD.CreateCell(3);//给第1行添加第1个单元格
            cell4GD.SetCellValue("合计");


            IRow rowHF = sheetHF.CreateRow(0);
            ICell cell1HF = rowHF.CreateCell(0);//给第1行添加第1个单元格
            cell1HF.SetCellValue("分类指标");
            ICell cell2HF = rowHF.CreateCell(1);//给第1行添加第1个单元格
            cell2HF.SetCellValue("指标分级");
            ICell cell3HF = rowHF.CreateCell(2);//给第1行添加第1个单元格
            cell3HF.SetCellValue("四川盆地");
            ICell cell4HF = rowHF.CreateCell(3);//给第1行添加第1个单元格
            cell4HF.SetCellValue("合计");

            int rowNumber = 1;
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

                    IRow rowGDSub = sheetGD.CreateRow(rowNumber);

                    ICell cell1 = rowGDSub.CreateCell(0);
                    if (j == 0)
                    {
                        cell1.SetCellValue(titleName);
                    }
                    else
                    {
                        cell1.SetCellValue("");
                    }
                    ICell cell2 = rowGDSub.CreateCell(1);
                    cell2.SetCellValue(itemJ.Key);

                    //ICellStyle cellStyle = workbook.CreateCellStyle();
                    //cellStyle.DataFormat =  new XSSFDataFormat().GetFormat("0.00");
                    

                    ICell cell3 = rowGDSub.CreateCell(2);
                    cell3.SetCellValue(itemJ.Value);
                    //cell3.CellStyle = cellStyle;

                    ICell cell4 = rowGDSub.CreateCell(3);
                    cell4.SetCellValue(itemJ.Value);
                    //cell4.CellStyle = cellStyle;

                    rowNumber++;
                }
            }

            if (File.Exists(filePathAndName))
            {
                File.Delete(filePathAndName);
            }

            using (FileStream file = new FileStream(filePathAndName, FileMode.Create))
            {
                workbook.Write(file);
            }

        }

    }
}
