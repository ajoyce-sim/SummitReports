using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using HtmlAgilityPack;
using IronPdf;
using SummitReports.Infrastructure;

namespace SummitReports.Objects
{
    public class DualSamplePdf : SummitPDFReportBaseObject, IGenericReport
    {
        DualSample _report;
        public DualSamplePdf() : base(@"Samples\DualSample\DualSamplePdf.html")
        {
            _report = new DualSample(this);
        }


        /// <summary>
        /// This will generate a RECommentPdf Report
        /// </summary>
        /// <param name="id">If this is zero, then we will assume that we are going top use BidPoolId</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int uwRECollateralId)
        {
            return await _report.GenerateAsync(uwRECollateralId);
        }
    }
}
