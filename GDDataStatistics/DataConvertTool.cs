using NPOI.SS.UserModel;

namespace GDDataStatistics
{
    public class DataConvertTool
    {
        /// <summary>
        /// 将cell值转化为真实需要的值 的数据类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public  static string getDealCellData(ICell cell)
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
