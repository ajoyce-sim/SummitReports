using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using SummitReports.Infrastructure;
using NPOI.SS.Converter;
using NPOI.XWPF.UserModel;
using System.Reflection;
using System;
using HtmlAgilityPack;

namespace SummitReports.Objects
{
    public class SummitWordReportBaseObject : SummitReportBaseObject<XWPFDocument>, ISummitReport, ISummitReportInternal
    {
        private NPoiWordDocument _document;
        public SummitWordReportBaseObject(string WordTemplatePathAndFileName) : base(WordTemplatePathAndFileName)
        {
            this.document = new XWPFDocument();
            this._document = new NPoiWordDocument(this.document);
        }
        public override ISummitDocument Document { get => _document; }

        protected override void ReadFile(string FileName)
        {

            using (FileStream file = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                this.document = new XWPFDocument(file);
                this._document = new NPoiWordDocument(this.document);
            }
        }

        protected override void WriteStream(FileStream file)
        {
            document.Write(file);
        }
    }
    public class TemplateExtentions
    {
        public TemplateExtentions()
        {

        }
        public TemplateExtentions(string fromExt, string toExt)
        {
            FromExtention = fromExt;
            ToExtention = toExt;
        }
        public string FromExtention { get; set; }
        public string ToExtention { get; set; }
    }
    public class SummitPDFReportBaseObject : SummitReportBaseObject<HtmlDocument>, ISummitReport, ISummitReportInternal
    { 
        private NPoiPdfDocument _document;
        public SummitPDFReportBaseObject(string WordTemplatePathAndFileName) : base(WordTemplatePathAndFileName)
        {
            _document = new NPoiPdfDocument(document);
            this.templateExtentions = new TemplateExtentions(".html", ".pdf");
        }
        public override ISummitDocument Document { get => _document;  }

        protected override void ReadFile(string FileName)
        {
            this.document = new HtmlDocument();
            this.document.Load(FileName);
            this._document = new NPoiPdfDocument(this.document);
        }

        protected override void WriteStream(FileStream file)
        {
            SelectPdf.HtmlToPdf converter = new SelectPdf.HtmlToPdf();
            SelectPdf.PdfDocument doc = converter.ConvertHtmlString(this.document.Text);
            doc.Save(file);
            doc.Close();
        }
    }
    public abstract class SummitReportBaseObject<T> : ISummitReport, ISummitReportInternal
    {
        public SummitReportBaseObject(string WordTemplatePathAndFileName)
        {
            this.ReportWorkPath = System.IO.Path.GetTempPath();
            var arr = WordTemplatePathAndFileName.Split('\\');
            wordTemplatePath = arr[0];
            wordTemplateFileName = arr[1];
        }

        protected T document;
        public virtual ISummitDocument Document { get => throw new Exception("Document not implemented"); set => throw new Exception("Document not implemented"); }
        public virtual string TemplateFileExtention { get => "docx";}

        protected string wordTemplateFileName = "";
        protected string wordTemplatePath = "";
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

        public virtual bool SaveToFile(string FileName)
        {
            using (var file2 = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite))
            {
                //document.Write(file2);
                WriteStream(file2);
                file2.Close();
            }
            return true;
        }

        protected abstract void WriteStream(FileStream file);
        protected abstract void ReadFile(string FileName);

        public void Clear()
        {
        }
        protected string templateFileName = "";
        protected string generatedFileName = "";
        public string GeneratedFileName { get => generatedFileName; set => generatedFileName = value; }
        public string TemplateFileName { get => templateFileName; set => templateFileName = value; }
        public TemplateExtentions templateExtentions = new TemplateExtentions(".docx", ".docx");

        public virtual bool ReloadTemplate(string initial = "")
        {
            this.GeneratedFileName = this.reportWorkPath + wordTemplateFileName.Replace(templateExtentions.FromExtention, "-" + Guid.NewGuid().ToString() + templateExtentions.ToExtention);

            var assembly = typeof(SummitReports.Objects.SummitExcelReportBaseObject).GetTypeInfo().Assembly;
            var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", wordTemplatePath, wordTemplateFileName));

            var lst2 = assembly.GetManifestResourceNames();

            try
            {
                FileStream fileStream = new FileStream(this.GeneratedFileName, FileMode.CreateNew);
                for (int i = 0; i < stream.Length; i++)
                    fileStream.WriteByte((byte)stream.ReadByte());
                fileStream.Close();
            }
            catch (Exception ex2)
            {
                throw new Exception(string.Format("Error while reading template {0}.{1} as an embedded resource, are you sure its spelled right and the you set the file Build Action as 'Embedded Resource'?", wordTemplatePath, wordTemplateFileName), ex2);
            }
            ReadFile(this.GeneratedFileName);
            return true;
        }
    }
}
