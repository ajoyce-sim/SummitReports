using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using SummitReports.Infrastructure;
using NPOI.SS.Converter;
using NPOI.XWPF.UserModel;

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
      
    }
}
