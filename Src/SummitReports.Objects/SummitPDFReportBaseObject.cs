using System.IO;
using SummitReports.Infrastructure;
//using IronPdf;
using HtmlAgilityPack;
using SelectPdf;
using System.Reflection;
using System;

namespace SummitReports.Objects
{

    public abstract class SummitPDFReportBaseObject : ISummitReport
    {
        public SummitPDFReportBaseObject(string PDFTemplatePathAndFileName)
        {
            this.ReportWorkPath = System.IO.Path.GetTempPath();
            var arr = PDFTemplatePathAndFileName.Split('\\');
            pdfTemplatePath = arr[0];
            pdfTemplateFileName = arr[1];
        }
        protected HtmlDocument document = null;

        public string TemplateFileName = "";
        public string GeneratedFileName = "";
        protected string pdfTemplateFileName = "";
        protected string pdfTemplatePath = "";
        protected string reportWorkPath;
        public string ReportWorkPath
        {
            get
            {
                return this.reportWorkPath;
            }
            set
            {
                this.reportWorkPath = value;
                if (!this.reportWorkPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
                {
                    this.reportWorkPath += Path.DirectorySeparatorChar.ToString();
                }
            }
        }

        public virtual bool ReloadTemplate(string initial = "")
        {
            try
            {
                this.GeneratedFileName = this.reportWorkPath + pdfTemplateFileName.Replace(".html", "-" + Guid.NewGuid().ToString() + ".pdf");

                var assembly = typeof(SummitReports.Objects.SummitExcelReportBaseObject).GetTypeInfo().Assembly;
                var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", pdfTemplatePath, pdfTemplateFileName));
                try
                {
                    FileStream fileStream = new FileStream(this.GeneratedFileName, FileMode.CreateNew);
                    for (int i = 0; i < stream.Length; i++)
                        fileStream.WriteByte((byte)stream.ReadByte());
                    fileStream.Close();
                }
                catch (Exception ex2)
                {
                    throw new Exception(string.Format("Error while reading template {0}.{1} as an embedded resource, are you sure its spelled right and the you set the file Build Action as 'Embedded Resource'?", pdfTemplatePath, pdfTemplateFileName), ex2);
                }
                using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
                {
                    this.document = new HtmlDocument();
                    this.document.Load(file);
                }
                return true;
            }
            catch (Exception)
            {

                throw;
            }

        }
        public bool SaveToFile(string FileName)
        {

            SelectPdf.HtmlToPdf converter = new SelectPdf.HtmlToPdf();
            SelectPdf.PdfDocument doc = converter.ConvertHtmlString(this.document.Text);
            doc.Save(FileName);
            doc.Close();
            /*
            HtmlToPdf converter = new HtmlToPdf();
            PdfDocument doc = converter.RenderHtmlAsPdf(this.document.Text);
            doc.SaveAs(FileName);
            */
            return true;
        }
        public void Clear()
        {
        }

    }
}
