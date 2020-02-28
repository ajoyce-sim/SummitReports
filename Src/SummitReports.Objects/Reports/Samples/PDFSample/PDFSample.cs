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
    public class PDFSample : SummitPDFReportBaseObject
    {
        public PDFSample() : base(@"Samples\PDFSample\PDFSample.html")
        {

        }


        /// <summary>
        /// This will generate a PDFSample Report
        /// </summary>
        /// <param name="id">If this is zero, then we will assume that we are going top use BidPoolId</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int uwRECollateralId)
        {
            try
            {
                if (!this.ReloadTemplate()) throw new Exception("Template could not be loaded :(");
                string sSQL = "";
                DataSet retDataSet = null;

                // Initialize Data Set
                sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_CollateralRE] WHERE [uwRECollateralId] = @p0;";

                retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL, uwRECollateralId);
                if ((retDataSet.Tables.Count == 1) && (retDataSet.Tables[0].Rows.Count == 1))
                {
                    var data = retDataSet.Tables[0].Rows[0];
                    Document.ReplaceFieldValue(data, "RptHeader");
                    Document.ReplaceFieldValue(data, "CollateralFullAddress");
                    Document.ReplaceFieldValue(data, "SIMValue", "C0");
                    Document.ReplaceFieldValue(data, "Comments");
                    SaveToFile(this.GeneratedFileName);
                    return this.GeneratedFileName;
                }
                return "No records found";
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
