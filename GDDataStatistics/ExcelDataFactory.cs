using GDDataStatistics.Model;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GDDataStatistics
{
    public class ExcelDataFactory
    {
        //列下标和列名对应关系
        private static Dictionary<int, string> cellIndexAndNameDic = new Dictionary<int, string>();
        //每张表统计出的结果-全量
        private static Dictionary<string, Dictionary<string, double>> totalDataDic = new Dictionary<string, Dictionary<string, double>>();

        //每张表统计的结果-按照行政区域统计
        private static ExcelDistrictDataInfo excelDistrictDataInfo = new ExcelDistrictDataInfo();


        //每张表统计的结果-按照行政区域统计
        private static Dictionary<string, Dictionary<string, Dictionary<string, double>>> totalDataDistrictDic = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();

        private static List<int> titleIndexList = new List<int>();
        private static List<string> titieNameList = new List<string>();

        /// <summary>
        /// 通过地址加载excel数据,分行政区统计
        /// </summary>
        /// <param name="filePath"></param>
        public static Dictionary<string, Dictionary<string, double>> LoadExcelDataDistrict(string filePath)
        {
            InitBasicData();

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var hssfworkbook = new XSSFWorkbook(file);
                var sheet = hssfworkbook.GetSheetAt(0);
                //循环行
                var rows = sheet.GetRowEnumerator();
                while (rows.MoveNext())
                {
                    var row = (XSSFRow)rows.Current;

                    if (row.RowNum <= 0) continue;
                    if (row.RowNum == 1)
                    {
                        //解析表列名
                        for (var i = 0; i < row.LastCellNum; i++)
                        {
                            var cell = row.GetCell(i);
                            if (cell == null)
                            {
                                continue;
                            }
                            else
                            {
                                string cellValue = DataConvertTool.getDealCellData(cell);
                                if (!string.IsNullOrWhiteSpace(cellValue) && titieNameList.Contains(cellValue))
                                {
                                    //需要统计的列下标
                                    titleIndexList.Add(i);
                                    cellIndexAndNameDic.Add(i, cellValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        //解析表数据
                        //图斑地类面积
                        double TBDLMJValue = 0;
                        //循环列
                        for (var i = 0; i < row.LastCellNum; i++)
                        {
                            if (!titleIndexList.Contains(i)) continue;

                            var cell = row.GetCell(i);
                            if (cell == null)
                            {
                                continue;
                            }
                            else
                            {
                                string cellValue = DataConvertTool.getDealCellData(cell);
                                string titleName = cellIndexAndNameDic[i];

                                if (titleName.Equals(TitleNameEnum.TBDLMJ.ToString()))
                                {
                                    TBDLMJValue = double.Parse(cellValue);
                                }
                                else
                                {
                                    //key为值相同的列，value为值相同列对应的图斑地类面积求和
                                    Dictionary<string, double> sameTypeDic = new Dictionary<string, double>();//每一列

                                    if (sameTypeDic.ContainsKey(cellValue))
                                    {
                                        sameTypeDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        sameTypeDic[cellValue] = TBDLMJValue;
                                    }

                                    if (totalDataDic.ContainsKey(titleName))
                                    {
                                        Dictionary<string, double> dicBase = totalDataDic[titleName];
                                        //合并两个dic;
                                        var dictionaries = new[] { dicBase, sameTypeDic };

                                        var mergeDicRes = DataMergeTool.MergeTwoDic(dictionaries);
                                        totalDataDic[titleName] = mergeDicRes;

                                    }
                                    else
                                    {
                                        totalDataDic[titleName] = sameTypeDic;

                                    }
                                }
                            }
                        }
                    }
                }
            }

            return totalDataDic;
        }


        /// <summary>
        /// 通过地址加载excel数据
        /// </summary>
        /// <param name="filePath"></param>
        public static Dictionary<string, Dictionary<string, double>> LoadExcelData(string filePath)
        {
            InitBasicData();

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var hssfworkbook = new XSSFWorkbook(file);
                var sheet = hssfworkbook.GetSheetAt(0);
                //循环行
                var rows = sheet.GetRowEnumerator();
                while (rows.MoveNext())
                {
                    var row = (XSSFRow)rows.Current;

                    if (row.RowNum <= 0) continue;
                    if (row.RowNum == 1)
                    {
                        //解析表列名
                        for (var i = 0; i < row.LastCellNum; i++)
                        {
                            var cell = row.GetCell(i);
                            if (cell == null)
                            {
                                continue;
                            }
                            else
                            {
                                string cellValue = DataConvertTool.getDealCellData(cell);
                                if (!string.IsNullOrWhiteSpace(cellValue) && titieNameList.Contains(cellValue))
                                {
                                    //需要统计的列下标
                                    titleIndexList.Add(i);
                                    cellIndexAndNameDic.Add(i, cellValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        //解析表数据
                        //图斑地类面积
                        double TBDLMJValue = 0;

                        //循环列
                        for (var i = 0; i < row.LastCellNum; i++)
                        {
                            if (!titleIndexList.Contains(i)) continue;

                            var cell = row.GetCell(i);
                            if (cell == null)
                            {
                                continue;
                            }
                            else
                            {
                                string cellValue = DataConvertTool.getDealCellData(cell);
                                string titleName = cellIndexAndNameDic[i];

                                if (titleName.Equals(TitleNameEnum.TBDLMJ.ToString()))
                                {
                                    TBDLMJValue = double.Parse(cellValue);
                                }
                                else if (string.Equals(titleName, TitleNameEnum.行政区代码.ToString(), StringComparison.OrdinalIgnoreCase) ||
                                   string.Equals(titleName, TitleNameEnum.行政区名称.ToString(), StringComparison.OrdinalIgnoreCase))
                                {
                                }
                                else
                                {
                                    //key为值相同的列，value为值相同列对应的图斑地类面积求和
                                    Dictionary<string, double> sameTypeDic = new Dictionary<string, double>();//每一列

                                    if (sameTypeDic.ContainsKey(cellValue))
                                    {
                                        sameTypeDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        sameTypeDic[cellValue] = TBDLMJValue;
                                    }

                                    if (totalDataDic.ContainsKey(titleName))
                                    {
                                        Dictionary<string, double> dicBase = totalDataDic[titleName];
                                        //合并两个dic;
                                        var dictionaries = new[] { dicBase, sameTypeDic };

                                        totalDataDic[titleName] = DataMergeTool.MergeTwoDic(dictionaries);
                                    }
                                    else
                                    {
                                        totalDataDic[titleName] = sameTypeDic;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return totalDataDic;
        }

        private static void InitBasicData()
        {
            titleIndexList.Clear();
            totalDataDic.Clear();
            cellIndexAndNameDic.Clear();
            totalDataDistrictDic.Clear();// 按照区域统计

            excelDistrictDataInfo = new ExcelDistrictDataInfo();

            Array names = Enum.GetNames(typeof(TitleNameEnum));
            foreach (var name in names)
            {
                string cellName = name.ToString();
                titieNameList.Add(cellName);
                if (!cellName.Equals(TitleNameEnum.TBDLMJ.ToString()))//此列是图斑面积
                {
                    totalDataDic.Add(cellName, new Dictionary<string, double>());
                }
            }
        }

    }
}
