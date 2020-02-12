using System;
using NPOI.SS.UserModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using System.Data;
using SummitReports.Infrastructure;
using NPOI.HSSF.UserModel;
using System.Collections.Generic;

namespace SummitReports.Objects
{
    public class AmortizationScheduleItem : IAmortizationScheduleItem
    {
        public int Month { get; set; }
        public DateTime ItemDate { get; set; }
        public decimal Factor { get; set; }
        public decimal BeginningBalance { get; set; }
        public decimal Payment { get; set; }
        public decimal Principal { get; set; }
        public decimal Interest { get; set; }
        public decimal Balloon { get; set; }
        public decimal EndBalance { get; set; }
    }
    public class AmortizationScheduleResult : List<IAmortizationScheduleItem>
    {

    }

    public class AmortizationCalculator : SummitReportBaseObject, IAmortizationCalculatorReport
    {
        public decimal UPB { get; set; } = 0.0m;
        public decimal CurrentlyPastDueInterest { get; set; } = 0.0m;
        public decimal CurrentlyPastDueFeesAndAdvances { get; set; } = 0.0m;
        public decimal PerDiemInterest { get; set; } = 0.0m;
        public decimal SIMFeesAndAdvances { get; set; } = 0.0m;
        public DateTime StartDate { get; set; } = DateTime.Now;
        public eInterestCalculationMethodology InterestCalculationMethodology { get; set; } = eInterestCalculationMethodology.Days365;
        public ePrincipalCalculation PrincipalCalculation { get; set; } = ePrincipalCalculation.Amortization;
        public int FixedPaymentAmount { get; set; } = 0;
        public decimal InterestRate { get; set; } = 0.0m;
        public int AmortizationTermYears { get; set; } = 0;
        public int BalloonPaymentMonths { get; set; } = 0;

        public AmortizationCalculator() : base(@"AmortizationCalculator\AmortizationCalculator_20200211.xlsx")
        {

        }

        List<IAmortizationScheduleItem> IAmortizationCalculatorReport.Calculate()
        {
            try
            {
                this.GeneratedFileName = this.reportWorkPath + excelTemplateFileName.Replace(".xlsx", "-" + Guid.NewGuid().ToString() + ".xlsx");

                var assembly = typeof(SummitReports.Objects.SummitReportBaseObject).GetTypeInfo().Assembly;
                var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", excelTemplatePath, excelTemplateFileName));
                FileStream fileStream = new FileStream(this.GeneratedFileName, FileMode.CreateNew);
                for (int i = 0; i < stream.Length; i++)
                    fileStream.WriteByte((byte)stream.ReadByte());
                fileStream.Close();
                using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
                {
                    this.workbook = new XSSFWorkbook(file);
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("Amortization Calculator"));
                }
                this.workbook.ClearStyleCache();
                /* VERY IMPORTANT NOTE - if you mess with the template,  sometimes you may have to Build->Clear Solution before the DLL that contains your template will get refreshed */

                List<string> Errors = new List<string>();
                if (this.UPB <= 0) Errors.Add(string.Format("UPB value of {0} is invalid.", this.UPB));
                if (this.StartDate.Equals(DateTime.MinValue)) Errors.Add(string.Format("StartDate value is not set."));
                if (this.FixedPaymentAmount <= 0) Errors.Add(string.Format("FixedPaymentAmount value of {0} is invalid.", this.FixedPaymentAmount));
                if (this.InterestRate <= 0) Errors.Add(string.Format("InterestRate value of {0} is invalid.", this.InterestRate));
                if (this.AmortizationTermYears <= 0) Errors.Add(string.Format("AmortizationTermYears value of {0} is invalid.", this.AmortizationTermYears));
                if (this.BalloonPaymentMonths <= 0) Errors.Add(string.Format("BalloonPaymentMonths value of {0} is invalid.", this.BalloonPaymentMonths));
                if (Errors.Count>0)
                {
                    throw new ArgumentException(string.Join(" ", Errors.ToArray()));
                }
                sheet.SetCellValue(2, "E", this, "UPB");
                sheet.SetCellValue(3, "E", this, "CurrentlyPastDueInterest");
                sheet.SetCellValue(4, "E", this, "CurrentlyPastDueFeesAndAdvances");
                sheet.SetCellValue(6, "E", this, "PerDiemInterest");
                sheet.SetCellValue(7, "E", this, "SIMFeesAndAdvances");
                sheet.SetCellValue(10, "E", this, "StartDate");
                sheet.SetCellValue(12, "E", this.InterestCalculationMethodology.ToString().Replace("Days", "Days "));
                sheet.SetCellValue(14, "E", this.PrincipalCalculation.ToString());
                sheet.SetCellValue(16, "E", this, "FixedPaymentAmount");
                sheet.SetCellValue(18, "E", this, "InterestRate");
                sheet.SetCellValue(19, "E", this, "AmortizationTermYears");
                sheet.SetCellValue(20, "E", this, "BalloonPaymentMonths");

                if (workbook is XSSFWorkbook)
                    XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
                else
                    HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
                var result = new AmortizationScheduleResult();
                var row = 23;
                for (int i = row; i < 500; i++)
                {
                    var endBalance = sheet.GetCellValue(i, "J", 0);
                    result.Add(new AmortizationScheduleItem()
                    {
                        Month = sheet.GetCellValue(i, "B", 0),
                        ItemDate = sheet.GetCellValue(i, "C", DateTime.MinValue),
                        Factor = sheet.GetCellValue(i, "D", 0.0m),
                        BeginningBalance = sheet.GetCellValue(i, "E", 0),
                        Payment = sheet.GetCellValue(i, "F", 0),
                        Principal = sheet.GetCellValue(i, "G", 0),
                        Interest = sheet.GetCellValue(i, "H", 0),
                        Balloon = sheet.GetCellValue(i, "I", 0),
                        EndBalance = endBalance
                    }) ;
                    if (endBalance == 0) break;
                }

                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
