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

    public class AmortizationCalculator : SummitExcelReportBaseObject, IAmortizationCalculatorReport
    {
        public decimal UPB { get; set; } = 0.0m;
        public decimal InterestRate { get; set; } = 0.0m;
        public eInterestCalculationMethodology InterestCalculationMethodology { get; set; } = eInterestCalculationMethodology.Days365;
        public int AmortizationTermYears { get; set; } = 0;
        public int BalloonPaymentMonths { get; set; } = 0;
        public DateTime StartDate { get; set; } = DateTime.Now;
        public int FirstPaymentMonth { get; set; } = 1;
        public decimal FixedPaymentAmount { get; set; } = 0.0m;
        private int paymentFrequency = 1;
        public int PaymentFrequency {
            get { return paymentFrequency; }
            set {
                var validValues = new List<int> { 1,2,4,6,12 };
                if (!validValues.Contains(value)) throw new ArgumentOutOfRangeException(string.Format("PaymentFrequency {0} is not 1,2,4,6,12", value));
                paymentFrequency = value;
            }
        }
        public bool IsInterestOnly { get; set; } = false;
        public int PIKStartMonth { get; set; } = 0;
        public int PIKEndMonth { get; set; } = 0;
        public int InterestOnlyEnd { get; set; } = 0;
        public bool IsFixedPayment { get; set; } = false;

        public AmortizationCalculator() : base(@"AmortizationCalculator\AmortizationCalculator_20200330.xlsx")
        {
        }

        AmortizationScheduleResult IAmortizationCalculatorReport.Calculate()
        {
            try
            {
                if (!this.ReloadTemplate("Sheet1")) throw new Exception("Template could not be loaded :(");

                List<string> Errors = new List<string>();
                if (this.UPB <= 0) Errors.Add(string.Format("UPB value of {0} is invalid.", this.UPB));
                if (this.StartDate.Equals(DateTime.MinValue)) Errors.Add(string.Format("StartDate value is not set."));
                if (this.FixedPaymentAmount < 0) Errors.Add(string.Format("FixedPaymentAmount value of {0} is invalid.", this.FixedPaymentAmount));
                if (this.InterestRate < 0) Errors.Add(string.Format("InterestRate value of {0} is invalid.", this.InterestRate));
                if (this.AmortizationTermYears <= 0) Errors.Add(string.Format("AmortizationTermYears value of {0} is invalid.", this.AmortizationTermYears));
                if (this.BalloonPaymentMonths < 0) Errors.Add(string.Format("BalloonPaymentMonths value of {0} is invalid.", this.BalloonPaymentMonths));
                if (this.InterestOnlyEnd < 0) Errors.Add(string.Format("InterestOnlyEnd value of {0} is invalid.", this.InterestOnlyEnd));
                if (this.PIKEndMonth < 0) Errors.Add(string.Format("PIKEndMonth value of {0} is invalid.", this.PIKEndMonth));
                if (this.PIKStartMonth < 0) Errors.Add(string.Format("PIKStartMonth value of {0} is invalid.", this.PIKStartMonth));
                if (Errors.Count>0)
                {
                    throw new ArgumentException(string.Join(" ", Errors.ToArray()));
                }
                
                sheet.SetCellValue(10, "F", this, "UPB");
                sheet.SetCellValue(12, "F", this, "InterestRate");
                if (double.TryParse(this.InterestCalculationMethodology.ToDescriptionString(), out var dblInterestCalculationMethodology))
                    sheet.SetCellValue(14, "F", dblInterestCalculationMethodology);
                sheet.SetCellValue(15, "F", this, "AmortizationTermYears");
                sheet.SetCellValue(16, "F", this, "BalloonPaymentMonths");
                sheet.SetCellValue(17, "F", this, "StartDate");
                sheet.SetCellValue(18, "F", this, "FirstPaymentMonth");
                sheet.SetCellValue(21, "F", this, "PaymentFrequency");
                sheet.SetCellValue(22, "F", ( this.IsInterestOnly ? "Y" : "N") );
                sheet.SetCellValue(23, "F", this, "PIKStartMonth");
                sheet.SetCellValue(24, "F", this, "PIKEndMonth");
                sheet.SetCellValue(25, "F", this, "InterestOnlyEnd");
                sheet.SetCellValue(26, "F", (this.IsFixedPayment ? "Y" : "N"));
                sheet.SetCellValue(27, "F", this, "FixedPaymentAmount");
                
                if (workbook is XSSFWorkbook)
                    XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
                
                else
                    HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);

                var result = new AmortizationScheduleResult();

                var row = 35;
                for (int i = row; i < 500; i++)
                {
                    var endBalance = sheet.GetCellValue(i, "J", 0.0m);
                    var payment = sheet.GetCellValue(i, "F", 0.0m);
                    //if (payment>0)
                    //{
                        result.AmortizationScheduleItemList.Add(new AmortizationScheduleItem()
                        {
                            Month = sheet.GetCellValue(i, "C", 0),
                            ItemDate = sheet.GetCellValue(i, "D", DateTime.MinValue),
                            Factor = 0,
                            BeginningBalance = sheet.GetCellValue(i, "E", 0.0m),
                            Payment = payment,
                            Principal = sheet.GetCellValue(i, "G", 0.0m),
                            Interest = sheet.GetCellValue(i, "H", 0.0m),
                            Balloon = sheet.GetCellValue(i, "I", 0.0m),
                            EndBalance = endBalance
                        });
                    //}
                    if (endBalance == 0) break;
                }
                result.GeneratedFileName = this.GeneratedFileName;
                return result;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                try
                {
                    SaveToFile(this.GeneratedFileName);
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
