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
    public class DocSample : SummitWordReportBaseObject
    {
        public DocSample() : base(@"Samples\DocSample\DocSample.docx")
        {

        }


        /// <summary>
        /// This will generate a PDFSample Report
        /// </summary>
        /// <param name="id">If this is zero, then we will assume that we are going top use BidPoolId</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int id)
        {
            try
            {
                if (!this.ReloadTemplate()) throw new Exception("Template could not be loaded :(");
                Document.ReplaceFieldValue("ENTERTEXT", string.Format("Id Passed was {0}", id));
                SaveToFile(this.GeneratedFileName);
                return this.GeneratedFileName;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
