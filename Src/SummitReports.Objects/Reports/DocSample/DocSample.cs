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
        public DocSample() : base(@"DocSample\DocSample.docx")
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
                foreach (var p in this.document.Paragraphs)
                {
                    if (p.ParagraphText.Contains("%ENTERTEXT%"))
                    {
                        p.ReplaceText("%ENTERTEXT%", string.Format("Id Passed was {0}", id));
                    }
                }
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
