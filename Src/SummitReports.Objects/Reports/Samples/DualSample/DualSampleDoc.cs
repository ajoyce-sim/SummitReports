using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using IronPdf;
using NPOI.XWPF.UserModel;
using SummitReports.Infrastructure;

namespace SummitReports.Objects
{
    public class DualSampleDoc : SummitWordReportBaseObject, IGenericReport
    {
        DualSample _report;
        public DualSampleDoc() : base(@"Samples\DualSampleDoc\DualSampleDoc.docx")
        {
            _report = new DualSample(this);
        }

        /// <summary>
        /// This will generate a Report with the Real Estate Comment.
        /// </summary>
        /// <param name="id">uwRECollateralId</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int uwRECollateralId)
        {
            return await _report.GenerateAsync(uwRECollateralId);
        }
    }
}
