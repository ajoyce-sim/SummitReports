using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using SummitReports.Infrastructure;
using NPOI.SS.Converter;
using NPOI.XWPF.UserModel;
using System.Reflection;
using System;

namespace SummitReports.Objects
{

    public abstract class SummitWordReportBaseObject : ISummitReport
    {
        public SummitWordReportBaseObject(string WordTemplatePathAndFileName)
        {
            this.ReportWorkPath = System.IO.Path.GetTempPath();
            var arr = WordTemplatePathAndFileName.Split('\\');
            wordTemplatePath = arr[0];
            wordTemplateFileName = arr[1];
        }
        protected XWPFDocument document = new XWPFDocument();

        public string TemplateFileName = "";
        public string GeneratedFileName = "";
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

        public bool SaveToFile(string FileName)
        {
            using (var file2 = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite))
            {
                document.Write(file2);
                file2.Close();
            }
            return true;
        }
        public void Clear()
        {
        }

        public virtual bool ReloadTemplate(string initial = "")
        {
            this.GeneratedFileName = this.reportWorkPath + wordTemplateFileName.Replace(".docx", "-" + Guid.NewGuid().ToString() + ".docx");

            var assembly = typeof(SummitReports.Objects.SummitExcelReportBaseObject).GetTypeInfo().Assembly;
            var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", wordTemplatePath, wordTemplateFileName));
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
            using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
            {
                this.document = new XWPFDocument(file);
            }
            return true;
        }
    }
}
