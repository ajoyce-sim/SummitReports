using Xunit;
using SummitReports.Objects;
using SummitReports.Infrastructure;

namespace SummitReports.Tests
{
    public class UWRelCashFlowGenTests
    {
        [Fact]
        public async void UWRelReportOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var uwrelcf = new UWRelationshipCashFlow();
            var generatedFileNameRel = await uwrelcf.RelationshipGenerateAsync(16);
            var generatedFileNameBidPool = await uwrelcf.BidPoolGenerateAsync(3);
        }

        [Fact]
        public async void ModelReportGenTestOk()
        {
            //SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            SummitReportSettings.Instance.ConnectionString = "data source=NSWIN10VM;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new ModelReport();
            var generatedFIleName = await rpt.GenerateAsync(2);
        }

        [Fact]

        public async void DeanSheetPresentationtGenTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new DeanSheetPresentation();
            var generatedFIleNameA = await rpt.GenerateAsync(31);
            var generatedFIleNameB = await rpt.GenerateAsync(46);
        }

        [Fact]

        public async void REReportPresTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new REReportPres();
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(37);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(268);
        }

        [Fact]

        public async void LoansPresTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new LoansReportPres();
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(10);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(51);
        }

        [Fact]
        public async void BAReportTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new BAReport();
            var generatedFIleName1 = await rpt.BidPoolGenerateAsync(10);
            var generatedFIleName2 = await rpt.RelationshipGenerateAsync(51);
        }

        [Fact]
        public async void UWRelationshipMasterReportOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new UWRelationshipMasterReport();
            var generatedFileName1 = await rpt.GenerateAsync(50);
            var generatedFileName2 = await rpt.GenerateAsync(51);
            var generatedFileName3 = await rpt.GenerateAsync(52);
        }       
    }
}

