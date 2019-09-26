using Xunit;
using SummitReports.Objects;
using SummitReport.Infrastructure;

namespace SummitReports.Tests
{
    public class InfrastructureTests
    {
        [Fact]
        public async void ReportLoaderTestOk()
        {
            var list = ReportLoader.Instance.ReportClassList();
            IGenericReport report = ReportLoader.Instance.CreateInstance<IGenericReport>("SummitReports.Objects.ModelReport");
        }


    }
 }

