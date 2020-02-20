using SummitReports.Infrastructure;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SummitReports.Infrastructure
{
    public enum eInterestCalculationMethodology
    {
        Days365,
        Days360
    }
    public enum ePrincipalCalculation
    {
        Amortization,
        InterestOnly,
        FixedPayment
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
        decimal CurrentlyPastDueInterest { get; set; }
        decimal CurrentlyPastDueFeesAndAdvances { get; set; }
        decimal PerDiemInterest { get; set; }
        decimal SIMFeesAndAdvances { get; set; }
        DateTime StartDate { get; set; }
        eInterestCalculationMethodology InterestCalculationMethodology { get; set; }
        ePrincipalCalculation PrincipalCalculation { get; set; }
        decimal FixedPaymentAmount { get; set; }
        decimal InterestRate { get; set; }
        int AmortizationTermYears { get; set; }
        int BalloonPaymentMonths { get; set; }

        List<IAmortizationScheduleItem> Calculate();
    }
}
