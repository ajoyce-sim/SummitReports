using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace SummitReports.Objects
{
    public static class Extentions
    {
        public static void SetCellValue(this ICell cell, decimal value)
        {
            cell.SetCellValue((decimal)value);
        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, DateTime value)
        {
            var row = worksheet.GetRow(rowPosition);
            var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
            cell.SetCellValue(value);
            cell.CellStyle.DataFormat = 14;
        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, double value)
        {
            // Get row
            var row = worksheet.GetRow(rowPosition);
            var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
            cell.SetCellValue(value);
        }
        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, decimal value)
        {
            // Get row
            var row = worksheet.GetRow(rowPosition);
            var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
            cell.SetCellValue(value);
        }

        public static void SetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, string value)
        {
            var row = worksheet.GetRow(rowPosition);
            var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
            cell.SetCellValue(value);
        }


        public static void SetCellFormula(this ISheet worksheet, int rowPosition, int columnPosition, string value)
        {
            var row = worksheet.GetRow(rowPosition);
            var cell = row.GetCell(columnPosition) ?? row.CreateCell(columnPosition);
            cell.SetCellFormula(value);
        }

        public static double GetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, double defaultValue)
        {
            // Get row
            var row = worksheet.GetRow(rowPosition);
            var cell = row.GetCell(columnPosition, MissingCellPolicy.RETURN_NULL_AND_BLANK);
            if (cell == null) return defaultValue;
            return cell.NumericCellValue;
        }

        public static decimal GetCellValue(this ISheet worksheet, int rowPosition, int columnPosition, decimal defaultValue)
        {
            // Get row
            var row = worksheet.GetRow(rowPosition);
            var cell = row.GetCell(columnPosition, MissingCellPolicy.RETURN_NULL_AND_BLANK);
            if (cell == null) return defaultValue;
            return (decimal)cell.NumericCellValue;
        }
    }
}
