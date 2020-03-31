using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using SummitReports.Infrastructure;
using NPOI.SS.Converter;
using System.Reflection;
using System;

namespace SummitReports.Objects
{

    public abstract class SummitExcelReportBaseObject : ISummitReport
    {
        public SummitExcelReportBaseObject(string ExcelTemplatePathAndFileName)
        {
            this.ReportWorkPath = System.IO.Path.GetTempPath();
            var arr = ExcelTemplatePathAndFileName.Split('\\');
            excelTemplatePath = arr[0];
            excelTemplateFileName = arr[1];
        }
        protected IWorkbook workbook = new XSSFWorkbook();
        protected ISheet sheet = new XSSFSheet();

        protected int rowIndex = 0;
        protected string templateFileName = "";
        protected string generatedFileName = "";
        public string TemplateFileName { get => templateFileName; set => templateFileName = value; }
        public string GeneratedFileName { get => generatedFileName; set => generatedFileName = value; }
        protected string excelTemplateFileName = "";
        protected string excelTemplatePath = "";
        protected string reportWorkPath = "";
        protected int iSheet = 1;

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
                workbook.Write(file2);
                file2.Close();
            }
            return true;
        }
        public void Clear()
        {
        }
        /// <summary>
        /// Reload from the excel template found in the resources
        /// </summary>
        /// <param name="initial">The First sheet to set sheet object to,it will use the name, but if start with an @ and a number, it will use that directly as the sheet index.  if this is empty it will assume the first sheet</param>
        /// <returns></returns>
        public virtual bool ReloadTemplate(string initial = "")
        {
            var extention = ".xlsx";
            if (extention.Equals(".xls")) workbook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            if (extention.Equals(".xlsx")) workbook = new XSSFWorkbook();

            this.GeneratedFileName = this.reportWorkPath + excelTemplateFileName.Replace(extention, "-" + Guid.NewGuid().ToString() + extention);
            var lst= System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames();

            var assembly = typeof(SummitReports.Objects.SummitExcelReportBaseObject).GetTypeInfo().Assembly;
            var lst2 = assembly.GetManifestResourceNames();
            var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", excelTemplatePath, excelTemplateFileName));
            try
            {
                FileStream fileStream = new FileStream(this.GeneratedFileName, FileMode.CreateNew);
                for (int i = 0; i < stream.Length; i++)
                    fileStream.WriteByte((byte)stream.ReadByte());
                fileStream.Close();
            }
            catch (Exception ex2)
            {
                throw new Exception(string.Format("Error while reading template {0}.{1} as an embedded resource, are you sure its spelled right and the you set the file Build Action as 'Embedded Resource'?", excelTemplatePath, excelTemplateFileName), ex2);
            }

            using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
            {
                //this.workbook = new XSSFWorkbook(file);
                if (extention.Equals(".xls")) workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(file);
                if (extention.Equals(".xlsx")) workbook = new XSSFWorkbook(file);

                if ((initial.StartsWith("@")) && (int.TryParse(initial.Replace("@", ""), out int initialIndex)))
                {
                    this.iSheet = initialIndex;
                    this.sheet = this.workbook.GetSheetAt(initialIndex);

                }
                else if (initial.Length>0)
                {
                    this.iSheet = this.workbook.GetSheetIndex(initial);
                    this.sheet = this.workbook.GetSheetAt(this.iSheet);
                }
                else
                {
                    this.iSheet = 0;
                    this.sheet = this.workbook.GetSheetAt(this.iSheet);
                }
            }
            this.workbook.ClearStyleCache();
            return true;
        }


        protected bool SaveAsHtml(string inputXlsFile)
        {

            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();

            // Set output parameters
            excelToHtmlConverter.OutputColumnHeaders = false;
            excelToHtmlConverter.OutputHiddenColumns = false;
            excelToHtmlConverter.OutputHiddenRows = false;
            excelToHtmlConverter.OutputLeadingSpacesAsNonBreaking = false;
            excelToHtmlConverter.OutputRowNumbers = false;
            excelToHtmlConverter.UseDivsToSpan = false;

            // Process the Excel file
            excelToHtmlConverter.ProcessWorkbook(workbook);

            // Output the HTML file
            excelToHtmlConverter.Document.Save(Path.ChangeExtension(inputXlsFile, "html"));
            return true;
        }
    }
}
