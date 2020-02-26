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
    public class RECommentDoc : SummitWordReportBaseObject
    {
        public RECommentDoc() : base(@"RECommentDoc\RECommentDoc.docx")
        {

        }


        /// <summary>
        /// This will generate a Report with the Real Estate Comment.
        /// </summary>
        /// <param name="uwRECollateralId">Real Estate Collateral Id</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int uwRECollateralId)
        {
            try
            {




                this.GeneratedFileName = this.reportWorkPath + wordTemplateFileName.Replace(".docx", "-" + Guid.NewGuid().ToString() + ".docx");

                var assembly = typeof(SummitReports.Objects.SummitExcelReportBaseObject).GetTypeInfo().Assembly;
                var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", wordTemplatePath, wordTemplateFileName));
                FileStream fileStream = new FileStream(this.GeneratedFileName, FileMode.CreateNew);
                for (int i = 0; i < stream.Length; i++)
                    fileStream.WriteByte((byte)stream.ReadByte());
                fileStream.Close();
                using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
                {
                    this.document = new XWPFDocument(file);
                }
                
                string sSQL = "";
                DataSet retDataSet = null;

                // Initialize Data Set
                sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_CollateralRE] WHERE [uwRECollateralId] = @p0;";

                retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL, uwRECollateralId);
                if ((retDataSet.Tables.Count == 1) && (retDataSet.Tables[0].Rows.Count == 1))
                {
                    var data = retDataSet.Tables[0].Rows[0];
                    foreach (var p in this.document.Paragraphs)
                    {
                        if (p.ParagraphText.Contains("%RptHeader%"))
                        {
                            p.ReplaceText("%RptHeader%", (string)data["RptHeader"]);
                        }
                        if (p.ParagraphText.Contains("%CollateralFullAddress%"))
                        {
                            p.ReplaceText("%CollateralFullAddress%", (string)data["CollateralFullAddress"]);
                        }
                        if (p.ParagraphText.Contains("%SIMValue%"))
                        {
                            p.ReplaceText("%SIMValue%", data.Value("SIMValue", "C0"));
                            //p.ReplaceText("%SIMValue%", string.Format("{0:C}",data["SIMValue"]));
                        }
                        if (p.ParagraphText.Contains("%Comments%"))
                        {
                            p.ReplaceText("%Comments%", data.Value("Comments"));
                        }
                    }
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
