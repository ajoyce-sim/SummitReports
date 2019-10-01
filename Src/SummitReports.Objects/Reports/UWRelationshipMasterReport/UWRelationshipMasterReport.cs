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

        void DoMerge(List<string> _sourceFiles, string targetFileName)
        {
            XSSFWorkbook target = new XSSFWorkbook();

            bool b = false;
            int xlsFileIdx = 0;
            foreach (string strFile in _sourceFiles)
            {
                xlsFileIdx++;
                XSSFWorkbook sourceXls = new XSSFWorkbook(strFile);
                if (xlsFileIdx == 1) target = new XSSFWorkbook();
                for (int i = 0; i < sourceXls.NumberOfSheets; i++)
                {
                    XSSFSheet sheet1 = sourceXls.GetSheetAt(i) as XSSFSheet;
                    sheet1.CopyTo(target, sheet1.SheetName+"_" + xlsFileIdx.ToString(), true, false);
                }
            }
            using (FileStream fs = new FileStream(targetFileName, FileMode.Create, FileAccess.Write))
            {
                target.Write(fs);
            }
        }

        /// <summary>
        /// Create a master report for one UWRelationship
        /// </summary>
        /// <param name="Id">UWRelationshipId</param>
        /// <returns></returns>
        public async Task<string> GenerateAsync(int Id)
        {

            this.GeneratedFileName = this.reportWorkPath + excelTemplateFileName.Replace(".xlsx", "-" + Guid.NewGuid().ToString() + ".xlsx");
            var cfReport = new UWRelationshipCashFlow();
            var loanReport = new LoansReportPres();
            var reReport = new REReportPres();
            var baReport = new BAReport();
            var reportList = new List<string>();
            reportList.Add(await cfReport.RelationshipGenerateAsync(Id));
            reportList.Add(await loanReport.RelationshipGenerateAsync(Id));
            reportList.Add(await reReport.RelationshipGenerateAsync(Id));
            reportList.Add(await baReport.RelationshipGenerateAsync(Id));
            DoMerge(reportList, this.GeneratedFileName); 
            return this.GeneratedFileName;
        }
    }
}
