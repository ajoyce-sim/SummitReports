using System;
using NPOI.SS.UserModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using System.Data;
using SummitReports.Infrastructure;
using System.Collections.Generic;

namespace SummitReports.Objects
{
    public class UWRelationshipMasterReport : SummitReportBaseObject, IGenericReport
    {
        public UWRelationshipMasterReport() : base(@"UWRelationshipMasterReport\UWRelationshipMasterReport.xlsx")
        {

        }

        IWorkbook DoMerge(Dictionary<string, string> _sourceFiles, string targetFileName)
        {
            XSSFWorkbook target = new XSSFWorkbook();

            bool b = false;
            int xlsFileIdx = 0;
            foreach (string strFile in _sourceFiles.Keys)
            {
                xlsFileIdx++;
                XSSFWorkbook sourceXls = new XSSFWorkbook(strFile);
                if (xlsFileIdx == 1) target = new XSSFWorkbook();
                for (int i = 0; i < sourceXls.NumberOfSheets; i++)
                {
                    XSSFSheet sheet1 = sourceXls.GetSheetAt(i) as XSSFSheet;
                    sheet1.CopyTo(target, _sourceFiles[strFile], true, true);
                }
            }
            return target;
        }

        /// <summary>
        /// Create a master report for one UWRelationship
        /// </summary>
        /// <param name="Id">UWRelationshipId</param>
        /// <returns></returns>
        public async Task<string> GenerateAsync(int Id)
        {

            this.GeneratedFileName = this.reportWorkPath + excelTemplateFileName.Replace(".xlsx", "-" + Guid.NewGuid().ToString() + ".xlsx");
            var cfReport = new UWRelationshipCashFlow() { ReportWorkPath = this.reportWorkPath };
            var loanReport = new LoansReportPres() { ReportWorkPath = this.reportWorkPath };
            var reReport = new REReportPres() { ReportWorkPath = this.reportWorkPath };
            var baReport = new BAReport() { ReportWorkPath = this.reportWorkPath };
            var reportListWithName = new Dictionary<string, string>();
            reportListWithName.Add(await cfReport.RelationshipGenerateAsync(Id), "Cash Flow");
            reportListWithName.Add(await loanReport.RelationshipGenerateAsync(Id), "Loan");
            reportListWithName.Add(await reReport.RelationshipGenerateAsync(Id), "Real Estate");
            reportListWithName.Add(await baReport.RelationshipGenerateAsync(Id), "Business Assets");
            var target = DoMerge(reportListWithName, this.GeneratedFileName);
            using (FileStream fs = new FileStream(this.GeneratedFileName, FileMode.Create, FileAccess.Write))
            {
                ((XSSFWorkbook)target).Write(fs);
            }
            return this.GeneratedFileName;
        }
    }
}
