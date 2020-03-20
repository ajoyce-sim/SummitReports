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
    public class SummitInvestorBookPdf : SummitPDFReportBaseObject, IGenericReport
    {
        public SummitInvestorBookPdf() : base(@"SummitInvestorBook\SummitInvestorBookPdf.html")
        {
        }
        /// <summary>
        /// This will generate a Report with the Real Estate Comment.
        /// </summary>
        /// <param name="uwRECollateralId">uwRECollateralId</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int uwRECollateralId)
        {
            try
            {
                if (!ReloadTemplate()) throw new Exception("Template could not be loaded :(");

                string sSQL = "";
                DataSet retDataSet = null;

                sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_CollateralRE] WHERE [uwRECollateralId] = @p0;";

                retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL, uwRECollateralId);
                if ((retDataSet.Tables.Count == 1) && (retDataSet.Tables[0].Rows.Count == 1))
                {
                    var data = retDataSet.Tables[0].Rows[0];
                    Document.ReplaceFieldValue(data, "RptHeader");
                    Document.ReplaceFieldValue(data, "OneLineAddress");
                    Document.ReplaceFieldValue(data, "SIMValue", "C0");
                    Document.ReplaceFieldValue(data, "Comments");
                    SaveToFile(GeneratedFileName);
                    return GeneratedFileName;
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
