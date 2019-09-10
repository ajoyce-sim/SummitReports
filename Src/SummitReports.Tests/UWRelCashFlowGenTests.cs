using System;
using Xunit;
using SummitReports.Objects;
using System.Collections.Generic;
using SummitReports.Objects.Services;
using System.Diagnostics;

namespace SummitReports.Tests
{
    public class UWRelCashFlowGenTests
    {
        [Fact]
        public async void UWRelDataOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var svc = new UWDataService();
            var reldata = await svc.FetchUWRelationshipData(13);
            var relCFdata = await svc.FetchUWRelationshipCashFlowsData(13);
        }

        [Fact]
        public async void UWRelReportOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var uwrelcf = new UWRelationshipCashFlow();
            var generatedFIleName = await uwrelcf.GenerateAsync(13);
            var generatedFIleName2 = await uwrelcf.GenerateAsync(12);
            var generatedFIleName3 = await uwrelcf.GenerateAsync(11);
        }

        [Fact]
        public async void ModelReportGenTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new ModelReport();
            var generatedFIleName = await rpt.GenerateAsync(13);
        }

        [Fact]
        public async void SampleReportGenTestOk()
        {
            SummitReportSettings.Instance.ConnectionString = "data source=summittest.database.windows.net;initial catalog=MARS;user=simsa;password=D3n^3r#$";
            var rpt = new SampleReport();
            var generatedFIleName = await rpt.GenerateAsync(2);
        }

    }
}
