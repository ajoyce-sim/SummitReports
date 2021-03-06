using Xunit;
using SummitReports.Objects;
using SummitReports.Infrastructure;

namespace SummitReports.Tests
{
    public class ReportTests
    {
        [Fact]
        public async void PDFReportOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new PDFSample();
            var genFileName = await rpt.GenerateAsync(5);
        }

        [Fact]
        public async void DocReportOk()
        {

            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new DocSample();
            var genFileName = await rpt.GenerateAsync(6670);
        }

        [Fact]
        public async void RECommentPdfOk()
        {

            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt2 = new RECommentPdf();
            var genFileName2 = await rpt2.GenerateAsync(727);
        }
    }
}

