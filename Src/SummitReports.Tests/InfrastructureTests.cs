using Xunit;
using SummitReports.Infrastructure;
using SummitReports.Objects;
using System;

namespace SummitReports.Tests
{
    public class InfrastructureTests
    {
        [Fact]
        public async void ReportLoaderTestOk()
        {
            var list = ReportLoader.Instance.ReportClassList();
            IGenericReport report = ReportLoader.Instance.CreateInstance<IGenericReport>("ModelReport");
        }

        [Fact]
        public async void UWRelReportOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var uwrelcf = ReportLoader.Instance.CreateInstance<IBidPoolRelationshipReport>("UWRelationshipCashFlow");
            var generatedFileNameRel = await uwrelcf.RelationshipGenerateAsync(13);
            var generatedFileNameBidPool = await uwrelcf.BidPoolGenerateAsync(2);
        }

        [Fact]
        public async void ModelReportGenTestOk()
        {
            //SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            SummitReportSettings.Instance.ConnectionString = "data source=NSWIN10VM;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = ReportLoader.Instance.CreateInstance<IGenericReport>("ModelReport");
            var generatedFIleName = await rpt.GenerateAsync(2);
        }

        [Fact]
        public async void DeanSheetPresentationtGenTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = ReportLoader.Instance.CreateInstance<IBidPoolReport>("DeanSheetPresentation");
            var generatedFIleName = await rpt.GenerateAsync(37);
        }

        [Fact]
        public async void REReportPresTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = ReportLoader.Instance.CreateInstance<IBidPoolRelationshipReport>("REReportPres");
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(37);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(296);
        }

        [Fact]
        public async void LoansPresTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = ReportLoader.Instance.CreateInstance<IBidPoolRelationshipReport>("LoansReportPres");
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(2);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(12);
        }

        [Fact]
        public async void BAReportTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = ReportLoader.Instance.CreateInstance<IBidPoolRelationshipReport>("BAReport");
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(2);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(12);
        }

        [Fact]
        public void AmortCalcTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = ReportLoader.Instance.CreateInstance<IAmortizationCalculatorReport>("AmortizationCalculator");
            rpt.UPB = 10000000;
            rpt.CurrentlyPastDueInterest = 0;
            rpt.CurrentlyPastDueFeesAndAdvances = 0;
            rpt.PerDiemInterest = 0;
            rpt.SIMFeesAndAdvances = 0;
            rpt.StartDate = DateTime.Parse("3/1/2019");
            rpt.InterestCalculationMethodology = eInterestCalculationMethodology.Days365;
            rpt.PrincipalCalculation = ePrincipalCalculation.Amortization;
            rpt.FixedPaymentAmount = 55000;
            rpt.InterestRate = 0.059m;
            rpt.AmortizationTermYears = 10;
            rpt.BalloonPaymentMonths = 48;
            var ret = rpt.Calculate();
            rpt.SaveToFile(@"C:\Temp\AmortTest.xlsx");
        }

    }
}

