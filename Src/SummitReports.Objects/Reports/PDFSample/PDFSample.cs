using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using IronPdf;
using SummitReports.Infrastructure;

namespace SummitReports.Objects
{
    public class PDFSample : SummitPDFReportBaseObject
    {
        /* Note that this sample page is NOT used in this example..... yet :) */
        public PDFSample() : base(@"PDFSample\PDFSample.pdf")
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
               return await Task.Run(() =>
                {
                    this.GeneratedFileName = this.reportWorkPath + pdfTemplateFileName.Replace(".pdf", "-" + Guid.NewGuid().ToString() + ".pdf");

                    var assembly = typeof(SummitReports.Objects.SummitExcelReportBaseObject).GetTypeInfo().Assembly;
                    var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", pdfTemplatePath, pdfTemplateFileName));
                    FileStream fileStream = new FileStream(this.GeneratedFileName, FileMode.CreateNew);
                    for (int i = 0; i < stream.Length; i++)
                        fileStream.WriteByte((byte)stream.ReadByte());
                    fileStream.Close();
                    using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
                    {
                        this.document = PdfDocument.FromFile(this.GeneratedFileName);
                    }

                    /* This will rewrite the pdf directly from HTML */

                    IronPdf.HtmlToPdf Renderer = new IronPdf.HtmlToPdf();
                    this.document = Renderer.RenderHtmlAsPdf(string.Format("<h1>Hello World<h1><h3>ID Passed was {0}</h3>", id));
                    SaveToFile(this.GeneratedFileName);
                    return this.GeneratedFileName;
                });
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
