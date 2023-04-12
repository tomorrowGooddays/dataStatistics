using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace GDDataStatistics
{
    public class ExcelDataFactory
    {
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
        private static Dictionary<string, double> BHQTRYJZ_1icDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQTRPHZJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQSWDYXJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQTRZJS_1Dic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQSZJBDic = new Dictionary<string, double>();
        private static Dictionary<string, double> BHQGDEJDLJDic = new Dictionary<string, double>();
        private static List<int> titleIndexList = new List<int>();
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
                    if (row.RowNum < 2) continue;

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
                            switch (i)
                            {
                                case (int)TitleNameEnum.TBDLMJ:
                                    TBDLMJValue = double.Parse(cellValue);
                                    break;
                                case (int)TitleNameEnum.ZRQDM:
                                    if (ZRQDMDic.ContainsKey(cellValue))
                                    {
                                        ZRQDMDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        ZRQDMDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.ZRQDM.ToString()] = ZRQDMDic;
                                    break;
                                case (int)TitleNameEnum.PDJB:
                                    if (PDJBDic.ContainsKey(cellValue))
                                    {
                                        PDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        PDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.PDJB.ToString()] = PDJBDic;
                                    break;
                                case (int)TitleNameEnum.TCHDJB:
                                    if (TCHDJBDic.ContainsKey(cellValue))
                                    {
                                        TCHDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TCHDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TCHDJB.ToString()] = TCHDJBDic;
                                    break;
                                case (int)TitleNameEnum.TRZDJB:
                                    if (TRZDJBDic.ContainsKey(cellValue))
                                    {
                                        TRZDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRZDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRZDJB.ToString()] = TRZDJBDic;
                                    break;
                                case (int)TitleNameEnum.TRYJZHLJB:
                                    if (TRYJZHLJBDic.ContainsKey(cellValue))
                                    {
                                        TRYJZHLJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRYJZHLJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRYJZHLJB.ToString()] = TRYJZHLJBDic;
                                    break;
                                case (int)TitleNameEnum.TRPHZJB:
                                    if (TRPHZJBDic.ContainsKey(cellValue))
                                    {
                                        TRPHZJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRPHZJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRPHZJB.ToString()] = TRPHZJBDic;
                                    break;
                                case (int)TitleNameEnum.SWDYXJB:
                                    if (SWDYXJBDic.ContainsKey(cellValue))
                                    {
                                        SWDYXJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        SWDYXJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.SWDYXJB.ToString()] = SWDYXJBDic;
                                    break;
                                case (int)TitleNameEnum.TRZJSWRJB:
                                    if (TRZJSWRJBDic.ContainsKey(cellValue))
                                    {
                                        TRZJSWRJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        TRZJSWRJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.TRZJSWRJB.ToString()] = TRZJSWRJBDic;
                                    break;
                                case (int)TitleNameEnum.SZJB:
                                    if (SZJBDic.ContainsKey(cellValue))
                                    {
                                        SZJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        SZJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.SZJB.ToString()] = SZJBDic;
                                    break;
                                case (int)TitleNameEnum.GDEJDLJB:
                                    if (GDEJDLJBDic.ContainsKey(cellValue))
                                    {
                                        GDEJDLJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        GDEJDLJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.GDEJDLJB.ToString()] = GDEJDLJBDic;
                                    break;
                                case (int)TitleNameEnum.BHQPDJB:

                                    if (BHQPDJBDic.ContainsKey(cellValue))
                                    {
                                        BHQPDJBDic[cellValue] += TBDLMJValue;
                                    }
                                    else
                                    {
                                        BHQPDJBDic.Add(cellValue, TBDLMJValue);
                                    }

                                    totalDataDic[TitleNameEnum.BHQPDJB.ToString()] = BHQPDJBDic;
                                    break;
                                case (int)TitleNameEnum.BHQTCHDJB:
                                    break;
                                case (int)TitleNameEnum.BHQTRZDJB:
                                    break;
                                case (int)TitleNameEnum.BHQTRYJZ_1:
                                    break;
                                case (int)TitleNameEnum.BHQTRPHZJB:
                                    break;
                                case (int)TitleNameEnum.BHQSWDYXJB:
                                    break;
                                case (int)TitleNameEnum.BHQTRZJS_1:
                                    break;
                                case (int)TitleNameEnum.BHQSZJB:
                                    break;
                                case (int)TitleNameEnum.BHQGDEJDLJ:
                                    break;
                                default:
                                    break;
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
            BHQTRYJZ_1icDic.Clear();
            BHQTRPHZJBDic.Clear();
            BHQSWDYXJBDic.Clear();
            BHQTRZJS_1Dic.Clear();
            BHQSZJBDic.Clear();
            BHQGDEJDLJDic.Clear();

            Array values = Enum.GetValues(typeof(TitleNameEnum));
            foreach (var value in values)
            {
                titleIndexList.Add((int)value);
            }

            totalDataDic.Add(TitleNameEnum.ZRQDM.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.PDJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.TCHDJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.TRZDJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.TRYJZHLJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.TRPHZJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.SWDYXJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.TRZJSWRJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.SZJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.GDEJDLJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQPDJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQTCHDJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQTRZDJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQTRYJZ_1.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQTRPHZJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQSWDYXJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQTRZJS_1.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQSZJB.ToString(), new Dictionary<string, double>());
            totalDataDic.Add(TitleNameEnum.BHQGDEJDLJ.ToString(), new Dictionary<string, double>());
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
