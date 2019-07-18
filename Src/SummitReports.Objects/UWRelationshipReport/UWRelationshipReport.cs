using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using NPOI.HPSF;
using System.IO;

namespace SummitReports.Objects
{
    public class NPV
    {
        public NPV()
        {
            SetupWorksheet();
        }
        private HSSFWorkbook workbook = new HSSFWorkbook();
        private ISheet sheet;
        private string rateFormula = "B1";
        public string RateFormula
        {
            get
            {
                return rateFormula;
            }
            set
            {
                rateFormula = value;
                this.sheet.SetCellFormula(0,3, string.Format("NPV({0},B4:B51)", rateFormula.Replace(NPV.RATE_VAR, "B1")));
            }
        }

        private int rowIndex = 0;
        public static readonly string RATE_VAR = "{RATE}";
        private void SetupWorksheet()
        {
            this.workbook = new HSSFWorkbook();
            this.sheet = workbook.CreateSheet("Financial");
            rowIndex = 0;
            var row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("Discount Rate");
            row.CreateCell(1).SetCellValue(0);
            row.CreateCell(2).SetCellValue("PV");
            row.CreateCell(3).SetCellFormula(string.Format("NPV({0},B4:B51)", RateFormula.Replace(NPV.RATE_VAR, "B1")));
            rowIndex++; rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("Period");
            row.CreateCell(1).SetCellValue("GCF");
            for (int i = 1; i < 100; i++)
            {
                rowIndex++;
                row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue(i);
                row.CreateCell(1).SetCellValue(0);
            }
        }
        public bool SaveToFile(string FileName)
        {
            using (var file2 = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite))
            {
                workbook.Write(file2);
                file2.Close();
            }
            return true;
        }
        public void Clear()
        {
            sheet.SetCellValue(0, 1, 0.0);
            for (int i = 3; i < 102; i++)
            {
                sheet.SetCellValue(i, 1, 0.0);
            }
        }
        public double Calculate(double rate, List<double> values)
        {
            this.Clear();
            sheet.SetCellValue(0, 1, rate);
            var startRow = 3;
            foreach (var value in values)
            {
                sheet.SetCellValue(startRow, 1, value);
                startRow++;
            }
            HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            var npvValue = sheet.GetCellValue(0, 3, 0d);
            return npvValue;
        }
    }
}
