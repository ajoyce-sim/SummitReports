﻿using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using SummitReports.Infrastructure;

namespace SummitReports.Objects
{

    public abstract class SummitReportBaseObject : ISummitReport
    {
        public SummitReportBaseObject(string ExcelTemplatePathAndFileName)
        {
            this.ReportWorkPath = System.IO.Path.GetTempPath();
            var arr = ExcelTemplatePathAndFileName.Split('\\');
            excelTemplatePath = arr[0];
            excelTemplateFileName = arr[1];
        }
        protected XSSFWorkbook workbook = new XSSFWorkbook();
        protected ISheet sheet;

        protected int rowIndex = 0;
        public string TemplateFileName = "";
        public string GeneratedFileName = "";
        protected string excelTemplateFileName = "";
        protected string excelTemplatePath = "";
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
                workbook.Write(file2);
                file2.Close();
            }
            return true;
        }
        public void Clear()
        {
        }

    }
}
