using FastMember;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace SummitReports.Objects
{
    public static class Extentions
    {
        public static void SetCellValue(this ICell cell, decimal value)
        {
            cell.SetCellValue((decimal)value);
        }
        public static bool isDateTime(this Object obj)
        {
            return (obj.GetType().Name.Contains("DateTime")) || (obj.GetType().UnderlyingSystemType.Name.Contains("DateTime"));
        }
        public static bool isDouble(this Object obj)
        {
            return (obj.GetType().Name.Contains("Int")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Int"))
                    || (obj.GetType().UnderlyingSystemType.Name.Contains("Decimal")) || (obj.GetType().UnderlyingSystemType.Name.Contains("double"))
                    || (obj.GetType().Name.Contains("Double")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Double"))
                    || (obj.GetType().Name.Contains("Single")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Single"))
                    || (obj.GetType().Name.Contains("Float")) || (obj.GetType().UnderlyingSystemType.Name.Contains("Float"))
                    ;
        }
        public static bool isBool(this Object obj)
        {
            return ((obj.GetType().Name.Contains("Bool")) || (obj.GetType().UnderlyingSystemType.Name.Contains("bool")));
        }
        public static bool isString(this Object obj)
        {
            return ((obj.GetType().Name.Contains("String")) || (obj.GetType().UnderlyingSystemType.Name.Contains("string")));
        }


        public static void SetCellValue<T>(this ISheet worksheet, int rowPosition, string columnLetter, T sourceObject, string FieldName)
        {
            try
            {
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;

                var accessor = TypeAccessor.Create(sourceObject.GetType());
                var obj = accessor[sourceObject, FieldName];
                if (obj == null)
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.Blank);
                    return;
                }
                if (obj.isDateTime())
                {
                    var c = TypeDescriptor.GetConverter(obj.GetType());
                    if (c.CanConvertTo(obj.GetType()))
                        worksheet.SetCellValue(rowPosition, columnPosition, (DateTime)c.ConvertTo(obj, DateTime.Now.GetType()));
                    else
                        worksheet.SetCellValue(rowPosition, columnPosition, (DateTime)obj);
                }
                else if (obj.isDouble())
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.Numeric);
                    var c = TypeDescriptor.GetConverter(obj.GetType());
                    if (c.CanConvertTo(obj.GetType()))
                        worksheet.SetCellValue(rowPosition, columnPosition, (double)c.ConvertTo(obj, 0.0.GetType()));
                    else 
                        worksheet.SetCellValue(rowPosition, columnPosition, Convert.ToDouble(obj));
                }
                else if (obj.isBool())
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.Boolean);
                    worksheet.SetCellValue(rowPosition, columnPosition, (bool)obj);
                    //var c = TypeDescriptor.GetConverter(obj.GetType());
                    //if (c.CanConvertTo(obj.GetType()))
                    //    worksheet.SetCellValue(rowPosition, columnPosition, (bool)c.ConvertTo(obj, true.GetType()));
                }
                else
                {
                    worksheet.SetCellType(rowPosition, columnPosition, CellType.String);
                    worksheet.SetCellValue(rowPosition, columnPosition, (string)obj);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1} for Field={2}.  {3}", rowPosition, columnLetter, FieldName, ex.Message));
            }

        }


        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, DateTime value)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                cell.CellStyle.DataFormat = 14;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, DateTime value)
        {
            try
            {
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
                cell.CellStyle.DataFormat = 14;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }

        public static void SetCellType(this ISheet worksheet, int rowPosition, int columnPosition, CellType type)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellType(type);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error Setting Cell Type of row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static void SetCellType(this ISheet worksheet, int rowPosition, string columnLetter, CellType type)
        {
            try
            {
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellType(type);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error Setting Cell Type of row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }


        public static void SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, double value)
        {
            try
            {
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, double value)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, bool value)
        {
            try
            {
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing bool value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }

        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, bool value)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing bool value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }
        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, decimal value)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }

        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, decimal value)
        {
            int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
            try
            {
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }
        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, string value)
        {
            try
            {
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnPosition, ex.Message));
            }
        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, string columnLetter, string value)
        {
            try
            {
                int columnPosition = columnLetter.ToCharArray().Select(c => c - 'A' + 1).Reverse().Select((v, i) => v * (int)Math.Pow(26, i)).Sum() - 1;
                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellValue(value);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(@"Error writing value to row={0} col={1}.  {2}", rowPosition, columnLetter, ex.Message));
            }
        }


        public static void SetCellFormula(this ISheet worksheet, int rowPosition, int columnPosition, string value)
        {
            try
            {

                var row = worksheet.GetRow(rowPosition - 1);
                var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
                cell.SetCellFormula(value);
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
                var row = worksheet.GetRow(rowPosition - 1);
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
                var row = worksheet.GetRow(rowPosition - 1);
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
                var row = worksheet.GetRow(rowPosition - 1);
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
                var row = worksheet.GetRow(rowPosition - 1);
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
