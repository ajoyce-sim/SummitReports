using SummitReports.Infrastructure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;

namespace SummitReports.Infrastructure
{
    public enum eInterestCalculationMethodology
    {
        [Description("365")]
        Days365,
        [Description("360")]
        Days360
    }

    public static class MyEnumExtensions
    {
        public static string ToDescriptionString(this eInterestCalculationMethodology val)
        {
            DescriptionAttribute[] attributes = (DescriptionAttribute[])val
               .GetType()
               .GetField(val.ToString())
               .GetCustomAttributes(typeof(DescriptionAttribute), false);
            return attributes.Length > 0 ? attributes[0].Description : string.Empty;
        }
    }

    public interface IAmortizationScheduleItem
    {
        int Month { get; set; }
        DateTime ItemDate { get; set; }
        decimal Factor { get; set; }
        decimal BeginningBalance { get; set; }
        decimal Payment { get; set; }
        decimal Principal { get; set; }
        decimal Interest { get; set; }
        decimal Balloon { get; set; }
        decimal EndBalance { get; set; }

    }
    public interface IAmortizationCalculatorReport : ISummitReport
    {
        decimal UPB { get; set; }
        decimal InterestRate { get; set; }
        eInterestCalculationMethodology InterestCalculationMethodology { get; set; }
        int AmortizationTermYears { get; set; }
        int BalloonPaymentMonths { get; set; }
        DateTime StartDate { get; set; }
        int FirstPaymentMonth { get; set; }
        int PaymentFrequency { get; set; }
        bool IsInterestOnly { get; set; }
        int PIKStartMonth { get; set; }
        int PIKEndMonth { get; set; }
        int InterestOnlyEnd { get; set; }
        bool IsFixedPayment { get; set; }
        decimal FixedPaymentAmount { get; set; }
        List<IAmortizationScheduleItem> Calculate();
    }
}
