using FastMember;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace SummitReports.Objects
{
    public static class NPoiExtentions
    {
        /// <summary>
        /// Protects SetSheetName with a try to we can figure out what name it died on.  It will also append the id at the end if the string becomes too large for the end of the sheet name (31 characters)
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static IWorkbook SetSheetName(this IWorkbook workbook, int sheetNumber,string sheetName, int id)
        {
            try
            {
                workbook.SetSheetName(sheetNumber, sheetName.AsSheetName(id));
                return workbook;

            }
            catch (Exception ex)
            {

                throw new Exception(string.Format("Could not rename sheet at {0} with name {0}. ", sheetNumber, sheetName), ex);
            }
        }

        /// <summary>
        /// Protects CreateSheet with a try to we can figure out what name it died on.  It will also append the id at the end if the string becomes too large for the end of the sheet name (31 characters)
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static ISheet CreateSheet(this IWorkbook workbook, string sheetName, int id)
        {
            try
            {
                return workbook.CreateSheet(sheetName.AsSheetName(id));
            }
            catch (Exception ex)
            {

                throw new Exception(string.Format("Could not create sheet with name {0}. ", sheetName), ex);
            }
        }
        public static void SetCellValue(this ICell cell, decimal value)
        {
            cell.SetCellValue((decimal)value);
        }
        public static ICell SetCellStyle(this ICell cell, ICellStyle style)
        {
            cell.CellStyle = style;
            return cell;
        }
        /// <summary>
        /// Special function that gives much more flexibility in applying style with different formats
        /// </summary>
        /// <param name="cell">Cell to apply this on</param>
        /// <param name="npoiStyle">NPoi Object previously created</param>
        /// <returns></returns>
        public static ICell SetCellStyle(this ICell cell, XSSFNPoiStyle npoiStyle)
        {
            cell.CellStyle = npoiStyle.Render(cell);
            return cell;
        }

        private static bool isDateTime(this Object obj)
        {
            return (obj.GetType().Name.Contains("DateTime")) || (obj.GetType().UnderlyingSystemType.Name.Contains("DateTime"));
        }
        private static bool isDouble(this Object obj)
        {
            return (obj.GetType().Name.Contains("Int")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Int"))
                    || (obj.GetType().UnderlyingSystemType.Name.Contains("Decimal")) || (obj.GetType().UnderlyingSystemType.Name.Contains("double"))
                    || (obj.GetType().Name.Contains("Double")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Double"))
                    || (obj.GetType().Name.Contains("Single")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Single"))
                    || (obj.GetType().Name.Contains("Float")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Float"))
                    ;
        }
        private static bool isBool(this Object obj)
        {
            return ((obj.GetType().Name.Contains("Bool")) || (obj.GetType().UnderlyingSystemType.Name.Contains("bool")));
        }
        private static bool isString(this Object obj)
        {
            return ((obj.GetType().Name.Contains("String")) || (obj.GetType().UnderlyingSystemType.Name.Contains("string")));
        }

        private static ICell GetCell(this ISheet worksheet, int rowPosition, int columnPosition)
        {
            var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
            return row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
        }
        private static ICell GetCell(this ISheet worksheet, int rowPosition, string columnLetter)
        {
            int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
            return worksheet.GetCell(rowPosition, columnPosition);
        }

        public static ICell SetCellValue<T>(this ISheet worksheet, int rowPosition, string columnLetter, T sourceObject, string FieldName, ICellStyle style)
        {
            return worksheet.SetCellValue(rowPosition, columnLetter, sourceObject, FieldName).SetCellStyle(style);
        }
        public static void ClearStyleCache(this IWorkbook workbook)
        {
            StyleCache.Clear();
        }
        public static Dictionary<string, ICellStyle> StyleCache = new Dictionary<string, ICellStyle>();
        /// <summary>
        /// This will create and cache a formula with scope of the worksheet.  but not that if you set use SetCellStyle,  that whatever is defined in SetCellStyle will be overridden by anything here
        /// </summary>
        /// <param name="cell">Exention this</param>
        /// <param name="formatString"></param>
        /// <returns>ICell - The Cell it is changing</returns>
        public static ICell SetCellFormat(this ICell cell, string formatString)
        {
            if (!StyleCache.ContainsKey(formatString))
            {
                var sheet = cell.Sheet.Workbook;
                ICellStyle cs = sheet.CreateCellStyle();
                cs.DataFormat = sheet.CreateDataFormat().GetFormat(formatString);
                StyleCache.Add(formatString, cs);
            }
            cell.CellStyle = StyleCache[formatString];
            return cell;
        }

        /// <summary>
        /// Special function that gives much more flexibility in applying style with different formats
        /// </summary>
        /// <param name="cell">Cell to apply this on</param>
        /// <param name="npoiStyle">NPoi Object previously created</param>
        /// <returns></returns>
        public static ICell SetCellFormat(this ICell cell, XSSFNPoiStyle npoiStyle)
        {
            cell.CellStyle = npoiStyle.Render(cell);
            return cell;
        }

        public static ICell SetCellFormula(this ICell cell, string formulaString)
        {
            cell.SetCellFormula(formulaString);
            return cell;
        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, System.Data.DataRow datarow, string FieldName)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;

                var obj = datarow[FieldName];
                if ((obj == null) || (datarow[FieldName].Equals(System.DBNull.Value)))
                {
                    return worksheet.SetCellType(rowPosition, columnPosition, CellType.Blank);
                }
                if (obj.isDateTime())
                {
                    var c = TypeDescriptor.GetConverter(obj.GetType());
                    if (c.CanConvertTo(obj.GetType()))
                        return worksheet.SetCellValue(rowPosition, columnPosition, (DateTime)c.ConvertTo(obj, DateTime.Now.GetType()));
                    else
                        return worksheet.SetCellValue(rowPosition, columnPosition, (DateTime)obj);
                }
                else if (obj.isDouble())
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.Numeric);
                    var c = TypeDescriptor.GetConverter(obj.GetType());
                    if (c.CanConvertTo(obj.GetType()))
                        return worksheet.SetCellValue(rowPosition, columnPosition, (double)c.ConvertTo(obj, 0.0.GetType()));
                    else
                        return worksheet.SetCellValue(rowPosition, columnPosition, Convert.ToDouble(obj));
                }
                else if (obj.isBool())
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.Boolean);
                    return worksheet.SetCellValue(rowPosition, columnPosition, (bool)obj);
                }
                else
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.String);
                    return worksheet.SetCellValue(rowPosition, columnPosition, (string)obj);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1} for Field={2}.  {3}", rowPosition, columnLetter, FieldName, ex.Message));
            }
        }

        public static ISheet SetColumnWidth(this ISheet thisWorksheet, string columnLetter, int width)
        {
            try
            {
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                thisWorksheet.SetColumnWidth(columnPosition, width);
                return thisWorksheet;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static ICell SetCellValue<T>(this ISheet worksheet, int rowPosition, string columnLetter, T sourceObject, string FieldName)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var accessor = TypeAccessor.Create(sourceObject.GetType());
                var obj = accessor[sourceObject, FieldName];
                if (obj == null)
                {
                    return worksheet.SetCellType(rowPosition, columnPosition, CellType.Blank);
                }
                if (obj.isDateTime())
                {
                    var c = TypeDescriptor.GetConverter(obj.GetType());
                    if (c.CanConvertTo(obj.GetType()))
                        return worksheet.SetCellValue(rowPosition, columnPosition, (DateTime)c.ConvertTo(obj, DateTime.Now.GetType()));
                    else
                        return worksheet.SetCellValue(rowPosition, columnPosition, (DateTime)obj);
                }
                else if (obj.isDouble())
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.Numeric);
                    var c = TypeDescriptor.GetConverter(obj.GetType());
                    if (c.CanConvertTo(obj.GetType()))
                        return worksheet.SetCellValue(rowPosition, columnPosition, (double)c.ConvertTo(obj, 0.0.GetType()));
                    else
                        return worksheet.SetCellValue(rowPosition, columnPosition, Convert.ToDouble(obj));
                }
                else if (obj.isBool())
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.Boolean);
                    return worksheet.SetCellValue(rowPosition, columnPosition, (bool)obj);
                }
                else
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.String);
                    return worksheet.SetCellValue(rowPosition, columnPosition, (string)obj);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1} for Field={2}.  {3}", rowPosition, columnLetter, FieldName, ex.Message));
            }

        }

        public static ISheet MergeCells(this ISheet worksheet, int firstRow, int lastRow, string firstColumn, string lastColumn)
        {
            if (worksheet.LastRowNum < lastRow) worksheet.CreateRow(lastRow);
            int iFirstColumn = firstColumn.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
            int iLastColumn = lastColumn.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;

            var cra1 = new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow, iFirstColumn, iLastColumn);
            worksheet.AddMergedRegion(cra1);
            return worksheet;

        }

        public static CellRangeAddress MergeCellsRange(this ISheet worksheet, int firstRow, int lastRow, string firstColumn, string lastColumn)
        {
            if (worksheet.LastRowNum < lastRow) worksheet.CreateRow(lastRow);
            int iFirstColumn = firstColumn.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
            int iLastColumn = lastColumn.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;

            var cra1 = new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow, iFirstColumn, iLastColumn);
            worksheet.AddMergedRegion(cra1);
            return cra1;

        }

        public static CellRangeAddress SetRangeStyle(this CellRangeAddress range, ISheet sheet, XSSFNPoiStyle npoiStyle)
        {
            npoiStyle.ApplyBorderToRange(sheet, range);
            return range;
        }
        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, DateTime value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                cell.CellStyle.DataFormat = 14;
                return cell;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, DateTime value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                cell.CellStyle.DataFormat = 14;
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }

        public static ICell SetCellType(this ISheet worksheet, int rowPosition, int columnPosition, CellType type)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellType(type);
                return cell;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error Setting Cell Type of row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static ICell SetCellType(this ISheet worksheet, int rowPosition, string columnLetter, CellType type)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellType(type);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error Setting Cell Type of row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }


        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, double value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, double value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, bool value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(((value) ? "Yes" : "No"));
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing bool value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, bool value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(((value) ? "Yes" : "No"));
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing bool value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }
        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, decimal value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, decimal value)
        {
            int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }
        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, string value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }
        }

        public static ICell SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, string value)
        {
            try
            {
                if (worksheet.LastRowNum < rowPosition) worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }
        }


        public static ICell SetCellFormula(this ISheet worksheet, int rowPosition, int columnPosition, string value)
        {
            try
            {

                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellFormula(value);
                return cell;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }
        }

        public static double GetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, double defaultValue)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null) return defaultValue;
                return cell.NumericCellValue;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error getting value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }
        }

        public static decimal GetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, decimal defaultValue)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                var cell = row.GetCell(columnPosition, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null) return defaultValue;
                return (decimal)cell.NumericCellValue;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error getting value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }
        }

        public static double GetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, double defaultValue)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var cell = row.GetCell(columnPosition, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null) return defaultValue;
                return cell.NumericCellValue;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error getting value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }
        }

        public static decimal GetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, decimal defaultValue)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var cell = row.GetCell(columnPosition, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null) return defaultValue;
                return (decimal)cell.NumericCellValue;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error getting value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }
    }
}