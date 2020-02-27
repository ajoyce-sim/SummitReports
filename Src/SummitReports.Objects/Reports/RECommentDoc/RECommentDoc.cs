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
    public class RECommentDoc : SummitWordReportBaseObject, IGenericReport
    {
        public RECommentDoc() : base(@"RECommentDoc\RECommentDoc.docx")
        {

        }


        /// <summary>
        /// This will generate a Report with the Real Estate Comment.
        /// </summary>
        /// <param name="id">uwRECollateralId</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int id)
        {
            try
            {
                if (!this.ReloadTemplate()) throw new Exception("Template could not be loaded :(");

                string sSQL = "";
                DataSet retDataSet = null;

                // Initialize Data Set
                sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_CollateralRE] WHERE [uwRECollateralId] = @p0;";

                retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL, id);
                if ((retDataSet.Tables.Count == 1) && (retDataSet.Tables[0].Rows.Count == 1))
                {
                    var data = retDataSet.Tables[0].Rows[0];
                    document.ReplaceFieldValue(data, "RptHeader");
                    document.ReplaceFieldValue(data, "CollateralFullAddress");
                    document.ReplaceFieldValue(data, "SIMValue", "C0");
                    document.ReplaceFieldValue(data, "Comments");
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
