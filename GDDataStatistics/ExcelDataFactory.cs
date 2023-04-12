using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace GDDataStatistics
{
    public class ExcelDataFactory
    {
        private static Dictionary<int, string> cellIndexAndNameDic = new Dictionary<int, string>();
        //每张表统计出的结果
        private static Dictionary<string, Dictionary<string, double>> totalDataDic = new Dictionary<string, Dictionary<string, double>>();

        private static Dictionary<string, double> ZRQDMDic = new Dictionary<string, double>();
        private static Dictionary<string, double> PDJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> TCHDJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> TRZDJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> TRYJZHLJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> TRPHZJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> SWDYXJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> TRZJSWRJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> SZJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> GDEJDLJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQPDJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQTCHDJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQTRZDJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQTRYJZ_1Dic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQTRPHZJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQSWDYXJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQTRZJS_1Dic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQSZJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQGDEJDLJDic = new Dictionary<string, double>();
        private static List<int> titleIndexList = new List<int>();
        private static List<string> titieNameList = new List<string>();
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
                                string cellValue = getDealCellData(cell);
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
                                string cellValue = getDealCellData(cell);
                                string titleName = cellIndexAndNameDic[i];

                                if (titleName.Equals(TitleNameEnum.TBDLMJ.ToString()))
                                {
                                    TBDLMJValue = double.Parse(cellValue);
                                }else if (titleName.Equals(TitleNameEnum.ZRQDM.ToString()))
                                {
                                    if (ZRQDMDic.ContainsKey(cellValue))
                                    {
                                        ZRQDMDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        ZRQDMDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.ZRQDM.ToString()] = ZRQDMDic;
                                }else if (titleName.Equals(TitleNameEnum.PDJB.ToString()))
                                {
                                    if (PDJBDic.ContainsKey(cellValue))
                                    {
                                        PDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        PDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.PDJB.ToString()] = PDJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.TCHDJB.ToString()))
                                {
                                    if (TCHDJBDic.ContainsKey(cellValue))
                                    {
                                        TCHDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TCHDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TCHDJB.ToString()] = TCHDJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.TRZDJB.ToString()))
                                {
                                    if (TRZDJBDic.ContainsKey(cellValue))
                                    {
                                        TRZDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRZDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRZDJB.ToString()] = TRZDJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.TRYJZHLJB.ToString()))
                                {
                                    if (TRYJZHLJBDic.ContainsKey(cellValue))
                                    {
                                        TRYJZHLJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRYJZHLJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRYJZHLJB.ToString()] = TRYJZHLJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.TRPHZJB.ToString()))
                                {
                                    if (TRPHZJBDic.ContainsKey(cellValue))
                                    {
                                        TRPHZJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRPHZJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRPHZJB.ToString()] = TRPHZJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.SWDYXJB.ToString()))
                                {
                                    if (SWDYXJBDic.ContainsKey(cellValue))
                                    {
                                        SWDYXJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        SWDYXJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.SWDYXJB.ToString()] = SWDYXJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.TRZJSWRJB.ToString()))
                                {
                                    if (TRZJSWRJBDic.ContainsKey(cellValue))
                                    {
                                        TRZJSWRJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRZJSWRJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRZJSWRJB.ToString()] = TRZJSWRJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.SZJB.ToString()))
                                {
                                    if (SZJBDic.ContainsKey(cellValue))
                                    {
                                        SZJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        SZJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.SZJB.ToString()] = SZJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.GDEJDLJB.ToString()))
                                {
                                    if (GDEJDLJBDic.ContainsKey(cellValue))
                                    {
                                        GDEJDLJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        GDEJDLJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.GDEJDLJB.ToString()] = GDEJDLJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQPDJB.ToString()))
                                {
                                    if (BHQPDJBDic.ContainsKey(cellValue))
                                    {
                                        BHQPDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQPDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQPDJB.ToString()] = BHQPDJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQTCHDJB.ToString()))
                                {
                                    if (BHQTCHDJBDic.ContainsKey(cellValue))
                                    {
                                        BHQTCHDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQTCHDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQTCHDJB.ToString()] = BHQTCHDJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQTRZDJB.ToString()))
                                {
                                    if (BHQTRZDJBDic.ContainsKey(cellValue))
                                    {
                                        BHQTRZDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQTRZDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQTRZDJB.ToString()] = BHQTRZDJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQTRYJZ_1.ToString()))
                                {
                                    if (BHQTRYJZ_1Dic.ContainsKey(cellValue))
                                    {
                                        BHQTRYJZ_1Dic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQTRYJZ_1Dic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQTRYJZ_1.ToString()] = BHQTRYJZ_1Dic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQTRPHZJB.ToString()))
                                {
                                    if (BHQTRPHZJBDic.ContainsKey(cellValue))
                                    {
                                        BHQTRPHZJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQTRPHZJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQTRPHZJB.ToString()] = BHQTRPHZJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQSWDYXJB.ToString()))
                                {
                                    if (BHQSWDYXJBDic.ContainsKey(cellValue))
                                    {
                                        BHQSWDYXJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQSWDYXJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQSWDYXJB.ToString()] = BHQSWDYXJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQTRZJS_1.ToString()))
                                {
                                    if (BHQTRZJS_1Dic.ContainsKey(cellValue))
                                    {
                                        BHQTRZJS_1Dic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQTRZJS_1Dic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQTRZJS_1.ToString()] = BHQTRZJS_1Dic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQSZJB.ToString()))
                                {
                                    if (BHQSZJBDic.ContainsKey(cellValue))
                                    {
                                        BHQSZJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQSZJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQSZJB.ToString()] = BHQSZJBDic;
                                }
                                else if (titleName.Equals(TitleNameEnum.BHQGDEJDLJ.ToString()))
                                {
                                    if (BHQGDEJDLJDic.ContainsKey(cellValue))
                                    {
                                        BHQGDEJDLJDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQGDEJDLJDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQGDEJDLJ.ToString()] = BHQGDEJDLJDic;
                                }
                                else
                                {
                                    //do nothing
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

            ZRQDMDic.Clear();
            PDJBDic.Clear();
            TCHDJBDic.Clear();
            TRZDJBDic.Clear();
            TRYJZHLJBDic.Clear();
            TRPHZJBDic.Clear();
            SWDYXJBDic.Clear();
            TRZJSWRJBDic.Clear();
            SZJBDic.Clear();
            GDEJDLJBDic.Clear();
            BHQPDJBDic.Clear();
            BHQTCHDJBDic.Clear();
            BHQTRZDJBDic.Clear();
            BHQTRYJZ_1Dic.Clear();
            BHQTRPHZJBDic.Clear();
            BHQSWDYXJBDic.Clear();
            BHQTRZJS_1Dic.Clear();
            BHQSZJBDic.Clear();
            BHQGDEJDLJDic.Clear();

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

        private static string getDealCellData(ICell cell)
        {
            string value = string.Empty;
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    value = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Numeric:
                case CellType.Formula:
                    try
                    {
                        value = cell.NumericCellValue.ToString();
                    }
                    catch
                    {
                        value = cell.StringCellValue;
                    }
                    break;
                case CellType.String:
                    value = cell.StringCellValue;
                    break;
                case CellType.Error:
                case CellType.Blank:
                    break;
                default:
                    value = cell.CellFormula;
                    break;
            }

            return value;
        }
    }
}
