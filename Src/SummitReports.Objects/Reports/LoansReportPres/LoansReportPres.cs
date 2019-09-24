using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using NPOI.HPSF;
using System.IO;
using System.Reflection;
using SummitReports.Objects.Services;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using System.Text;
using System.Data.SqlClient;

namespace SummitReports.Objects
{
    public class LoansReportPres : SummitReportBaseObject
    {
        public LoansReportPres() : base(@"LoansReportPres\LoansReportPres.xlsx")
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="BidPoolId"></param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int BidPoolId)
        {
            try
            {
                // Initialize the workbook

                this.GeneratedFileName = this.reportWorkPath + excelTemplateFileName.Replace(".xlsx", "-" + Guid.NewGuid().ToString() + ".xlsx");

                var assembly = typeof(SummitReports.Objects.SummitReportSettings).GetTypeInfo().Assembly;
                var stream = assembly.GetManifestResourceStream(string.Format("SummitReports.Objects.Reports.{0}.{1}", excelTemplatePath, excelTemplateFileName));
                FileStream fileStream = new FileStream(this.GeneratedFileName, FileMode.CreateNew);
                for (int i = 0; i < stream.Length; i++)
                    fileStream.WriteByte((byte)stream.ReadByte());
                fileStream.Close();
                var iSheet = 1;
                using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
                {
                    this.workbook = new XSSFWorkbook(file);
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex(iSheet.ToString()));
                }
                this.workbook.ClearStyleCache();

                // Generate a Sheet for each relationship.

                string sSQL1 = @"SET ANSI_WARNINGS OFF; SELECT COUNT(*) AS TabCnt FROM UW.tbl_Relationship WHERE BidPoolId =@p0;";
                var retTabCnt = await MarsDb.QueryAsDataSetAsync(sSQL1, BidPoolId);
                System.Data.DataTable aResultSet = retTabCnt.Tables[0];
                var iTabCnt = 0;
                foreach (System.Data.DataRow a in aResultSet.Rows)
                { iTabCnt = (int)a["TabCnt"]; }

                for (int x = 2; x < iTabCnt + 1; x++)
                {
                    sheet = workbook.CloneSheet(this.workbook.GetSheetIndex("1"));
                    workbook.SetSheetName(workbook.NumberOfSheets - 1, x.ToString());
                }
                
                // Return to sheet "1"
                this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex(iSheet.ToString()));

                // Get Dataset for report using ADO 
                string sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_Loans] WHERE [BidPoolId]=@p0 ORDER BY uwRelationshipId ASC, uwLoanId ASC;";
                var retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, BidPoolId);
                System.Data.DataTable firstResultSet = retDataSet.Tables[0];
                var iRow = 1;
                var iRel = 0;
                var iLnCnt = 1;
                foreach (System.Data.DataRow row in firstResultSet.Rows)
                {
                    
                    if (iRow == 1)
                    {
                        iRel = (int) row["uwRelationshipId"];
                    }
                    else if (iRel != (int)row["uwRelationshipId"])
                    {
                        iSheet++;
                        this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex(iSheet.ToString()));
                        iRow = 1;
                        iLnCnt = 1;
                        iRel = (int)row["uwRelationshipId"];
                    }

                    var formatStr = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                    var LnCellStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, CellFormat = formatStr, VerticalAlignment = VerticalAlignment.Top, HorizontalAlignment = HorizontalAlignment.Left };

                    
                    sheet.SetCellValue(2, "B", row, "RptHeader");
                    sheet.CreateRow(iRow + 5);
                    
                    sheet.SetCellValue(iRow + 5, "B", row, "BidSubPoolName").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "C", row, "RelationshipName").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "D", row, "LoanShortName").SetCellStyle(LnCellStyle);
                    LnCellStyle.WrapText = true;
                    sheet.SetCellValue(iRow + 5, "E", row, "LoanDescriptionTxt").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "F", row, "BorrowerTxt").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "G", row, "GuarantorTxt").SetCellStyle(LnCellStyle);
                    LnCellStyle.CellFormat = "mm/dd/yyy";
                    sheet.SetCellValue(iRow + 5, "H", row, "OriginationDate").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "I", row, "MaturityDate").SetCellStyle(LnCellStyle);
                    LnCellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 5, "J", row, "OriginalUPB").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "K", row, "UPB").SetCellStyle(LnCellStyle);
                    LnCellStyle.CellFormat = "0.0%";
                    sheet.SetCellValue(iRow + 5, "L", row, "InterestRate").SetCellStyle(LnCellStyle);
                    LnCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 5, "M", row, "SIMValueLoan").SetCellStyle(LnCellStyle);

                    if (iLnCnt == (int)row["LoansCnt"])
                    {
                        //sheet.CreateRow(18 + iRow);
                        //sheet.SetCellValue(18 + iRow, "C", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(C18:C{0})", (18 + iRow - 2)));
                        sheet.CreateRow(iRow + 7);
                        LnCellStyle.IsBold = true; 
                        sheet.SetCellValue(iRow + 6, "C", "Totals:").SetCellStyle(LnCellStyle);
                        LnCellStyle.CellFormat = "#,##0.00";
                        sheet.SetCellValue(iRow + 6, "J", 0.0).SetCellStyle(LnCellStyle).SetCellFormula(string.Format("SUM(J6:J{0})", (6 + iRow )));
                        sheet.SetCellValue(iRow + 6, "K", 0.0).SetCellStyle(LnCellStyle).SetCellFormula(string.Format("SUM(K6:K{0})", (6 + iRow )));
                        LnCellStyle.CellFormat = "#,###";
                        sheet.SetCellValue(iRow + 6, "M", 0.0).SetCellStyle(LnCellStyle).SetCellFormula(string.Format("SUM(M6:M{0})", (6 + iRow )));
                        LnCellStyle.IsBold = false;
                    }

                    iRow++;
                    iLnCnt++;

                }

                SaveToFile(this.GeneratedFileName);
                return this.GeneratedFileName;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
