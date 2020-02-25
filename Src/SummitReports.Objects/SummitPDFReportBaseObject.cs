using System.IO;
using SummitReports.Infrastructure;
using IronPdf;

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
        protected PdfDocument document = null;

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

        public bool SaveToFile(string FileName)
        {
            this.document.SaveAs(FileName);
            return true;
        }
        public void Clear()
        {
        }
      
    }
}
