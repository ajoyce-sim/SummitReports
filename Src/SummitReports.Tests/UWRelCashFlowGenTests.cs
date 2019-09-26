using Xunit;
using SummitReports.Objects;

namespace SummitReports.Tests
{
    public class UWRelCashFlowGenTests
    {
        [Fact]
        public async void UWRelReportOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var uwrelcf = UWRelationshipCashFlow.CreateInstance();
            var generatedFileNameRel = await uwrelcf.RelationshipGenerateAsync(13);
            var generatedFileNameBidPool = await uwrelcf.BidPoolGenerateAsync(2);
        }

        [Fact]
        public async void ModelReportGenTestOk()
        {
            //SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            SummitReportSettings.Instance.ConnectionString = "data source=NSWIN10VM;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = ModelReport.CreateInstance();
            var generatedFIleName = await rpt.GenerateAsync(2);
        }

        [Fact]

        public async void DeanSheetPresentationtGenTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = DeanSheetPresentation.CreateInstance();
            var generatedFIleName = await rpt.GenerateAsync(37);
        }

        [Fact]

        public async void REReportPresTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = REReportPres.CreateInstance();
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(37);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(268);
        }

        [Fact]

        public async void LoansPresTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = LoansReportPres.CreateInstance();
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(2);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(12);
        }

        [Fact]


        public async void BAReportTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new BAReport();
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(2);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(12);
        }
    }

 }

