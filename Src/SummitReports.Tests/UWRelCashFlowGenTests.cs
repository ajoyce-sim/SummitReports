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
            var generatedFIleName = await rpt.GenerateAsync(37);
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

