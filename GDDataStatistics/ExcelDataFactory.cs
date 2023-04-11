using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;

namespace GDDataStatistics
{
    public class ExcelDataFactory
    {
        /// <summary>
        /// 通过地址加载excel数据
        /// </summary>
        /// <param name="filePath"></param>
        public static DataTable LoadExcelData(string filePath)
        {
            var dt = new DataTable();
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var hssfworkbook = new XSSFWorkbook(file);
                var sheet = hssfworkbook.GetSheetAt(0);
                //for (var j = 0; j < 5; j++)
                //{
                //    dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
                //}
                var rows = sheet.GetRowEnumerator();
                while (rows.MoveNext())
                {
                    var row = (XSSFRow)rows.Current;
                    var dr = dt.NewRow();
                    for (var i = 0; i < row.LastCellNum; i++)
                    {
                        var cell = row.GetCell(i);
                        if (cell == null)
                        {
                            dr[i] = null;
                        }
                        else
                        {
                            switch (cell.CellType)
                            {
                                case CellType.Blank:
                                    dr[i] = "[null]";
                                    break;
                                case CellType.Boolean:
                                    dr[i] = cell.BooleanCellValue;
                                    break;
                                case CellType.Numeric:
                                    dr[i] = cell.ToString();
                                    break;
                                case CellType.String:
                                    dr[i] = cell.StringCellValue;
                                    break;
                                case CellType.Formula:
                                    try
                                    {
                                        dr[i] = cell.NumericCellValue;
                                    }
                                    catch
                                    {
                                        dr[i] = cell.StringCellValue;
                                    }
                                    break;
                                case CellType.Error:
                                    break;
                                default:
                                    dr[i] = "=" + cell.CellFormula;
                                    break;
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
    }
}
