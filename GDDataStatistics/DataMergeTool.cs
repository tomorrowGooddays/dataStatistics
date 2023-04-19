using GDDataStatistics.Helper;
using GDDataStatistics.Model;
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


        public static string ExportData(string filePathAndName, Dictionary<string, Dictionary<string, double>> dicData)
        {
            if (string.IsNullOrEmpty(filePathAndName)) return "";

            filePathAndName = filePathAndName.Replace(".xlsx", $"{DateTime.Now.ToString("yyyyMMdd")}.xlsx");

            //创建workbook，说白了就是在内存中创建一个Excel文件
            IWorkbook workbook = CreateWorkbook();
            ISheet sheetGD = workbook.GetSheet(SheetNameEnum.GD.ToString());

            // 创建样式对象
            var style = workbook.CreateCellStyle();
            // 设置单元格格式为数字格式，并保留两位小数
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");

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

                    ICell cell3 = rowGDSub.CreateCell(2);
                    cell3.SetCellValue(itemJ.Value);
                    cell3.CellStyle = style;

                    ICell cell4 = rowGDSub.CreateCell(3);
                    cell4.SetCellValue(itemJ.Value);
                    cell4.CellStyle = style;

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


            return filePathAndName;
        }

        private static IWorkbook CreateWorkbook()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheetGD = workbook.CreateSheet(SheetNameEnum.GD.ToString());
            ISheet sheetHF = workbook.CreateSheet(SheetNameEnum.HF.ToString());

            // 创建样式对象
            var style = workbook.CreateCellStyle();
            // 设置单元格格式为数字格式，并保留两位小数
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");

            IRow rowGD = sheetGD.CreateRow(0);//添加第1行,注意行列的索引都是从0开始的

            ICell cell1GD = rowGD.CreateCell(0);//给第1行添加第1个单元格
            cell1GD.SetCellValue("分类指标");
            ICell cell2GD = rowGD.CreateCell(1);
            cell2GD.SetCellValue("指标分级");
            ICell cell3GD = rowGD.CreateCell(2);
            cell3GD.SetCellValue("四川盆地");
            ICell cell4GD = rowGD.CreateCell(3);
            cell4GD.SetCellValue("合计");


            IRow rowHF = sheetHF.CreateRow(0);
            ICell cell1HF = rowHF.CreateCell(0);
            cell1HF.SetCellValue("分类指标");
            ICell cell2HF = rowHF.CreateCell(1);
            cell2HF.SetCellValue("指标分级");
            ICell cell3HF = rowHF.CreateCell(2);
            cell3HF.SetCellValue("四川盆地");
            ICell cell4HF = rowHF.CreateCell(3);
            cell4HF.SetCellValue("合计");

            return workbook;
        }

        public static void ExportDataByName(string filePath, List<ExcelDataInfo> dataList)
        {
            string tableNameMapJsonFileName = "tableNameMap.json";
            List<TableNameMap> tableNameMaps = FileHelper.GetJsonFileFromEmbedResource<List<TableNameMap>>(tableNameMapJsonFileName);

            if (tableNameMaps == null || tableNameMaps.Count == 0)
            {
                throw new Exception("tableNameMap.json文件数据缺失");
            }

            foreach (var tableNameMap in tableNameMaps)
            {
                string excelName = tableNameMap.CnName;

                var excelInfo = tableNameMap.ExcelInfo;

                List<string> enNameList = tableNameMap.ExcelInfo.Select(p => p.EnName).ToList();

                List<ExcelDataInfo> dataInfos = dataList.Where(x => enNameList.Any(y => string.Equals(x.FileName, y, StringComparison.OrdinalIgnoreCase))).ToList();

                if (dataInfos != null && dataInfos.Count > 0)
                {
                    DoExport(filePath, tableNameMap, dataInfos);
                }

            }

        }

        private static void DoExport(string filePath, TableNameMap tableNameMap, List<ExcelDataInfo> dataInfos)
        {
            string fileDirectory = $"{filePath}\\成果输出";
            if (!Directory.Exists(fileDirectory))
            {
                Directory.CreateDirectory(fileDirectory);
            }

            string filePathAndName = $"{fileDirectory}\\{tableNameMap.CnName}.xlsx";
            //创建好表头
            IWorkbook workbook = new XSSFWorkbook();

            // 创建样式对象
            var style = workbook.CreateCellStyle();
            // 设置单元格格式为数字格式，并保留两位小数
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");

            for (int s = 0; s < dataInfos.Count; s++)
            {
                string sheetName = tableNameMap.ExcelInfo.FirstOrDefault(p => string.Equals(p.EnName, dataInfos[s].FileName, StringComparison.OrdinalIgnoreCase))?.SheetName;
                if (string.IsNullOrWhiteSpace(sheetName)) sheetName = SheetNameEnum.GD.ToString();

                ISheet sheet = workbook.CreateSheet(sheetName);

                IRow row = sheet.CreateRow(0);//添加第1行,注意行列的索引都是从0开始的

                ICell cell1 = row.CreateCell(0);//给第1行添加第1个单元格
                cell1.SetCellValue("分类指标");
                ICell cell2 = row.CreateCell(1);
                cell2.SetCellValue("指标分级");
                ICell cell3 = row.CreateCell(2);
                cell3.SetCellValue("四川盆地");
                ICell cell4 = row.CreateCell(3);
                cell4.SetCellValue("合计");

                Dictionary<string, Dictionary<string, double>> dicData = dataInfos[s].DataList;

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

                        IRow rowGDSub = sheet.CreateRow(rowNumber);

                        ICell cell1J = rowGDSub.CreateCell(0);
                        if (j == 0)
                        {
                            cell1J.SetCellValue(titleName);
                        }
                        else
                        {
                            cell1J.SetCellValue("");
                        }
                        ICell cell2J = rowGDSub.CreateCell(1);
                        cell2J.SetCellValue(itemJ.Key);

                        ICell cell3J = rowGDSub.CreateCell(2);
                        cell3J.SetCellValue(itemJ.Value);
                        cell3J.CellStyle = style;

                        ICell cell4J = rowGDSub.CreateCell(3);
                        cell4J.SetCellValue(itemJ.Value);
                        cell4J.CellStyle = style;

                        rowNumber++;
                    }
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
