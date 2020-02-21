using System;
using System.Data;
using NPOI.SS.UserModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using SummitReports.Infrastructure;

namespace SummitReports.Objects
{
    public class LoansReportPres : SummitReportBaseObject, IBidPoolRelationshipReport
    {
        public LoansReportPres() : base(@"LoansReportPres\LoansReportPres.xlsx")
        {

        }

        /// <summary>
        /// This will generate a Loan report for all relationships in a bid pool (one relationship per tab)
        /// </summary>
        /// <param name="BidPoolId"></param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> BidPoolGenerateAsync(int BidPoolId)
        {
            return await GenerateAsync(BidPoolId, 0);
        }

        /// <summary>
        /// This will generate a Loan report for one Relationship (one tab)
        /// </summary>
        /// <param name="uwRelationshipId"></param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> RelationshipGenerateAsync(int uwRelationshipId)
        {
            return await GenerateAsync(0, uwRelationshipId);
        }

        /// <summary>
        /// This will generate a Business Asset Report for a BidPool or Relationship
        /// </summary>
        /// <param name="BidPoolId">If this is 0, then we will assume that we are going to use uwRelationshipId</param>
        /// <param name="uwRelationshipId">If this is zero, then we will assume that we are going top use BidPoolId</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int BidPoolId, int uwRelationshipId)
        {
            try
            {
                // Initialize the workbook

                this.GeneratedFileName = this.reportWorkPath + excelTemplateFileName.Replace(".xlsx", "-" + Guid.NewGuid().ToString() + ".xlsx");

                var assembly = typeof(SummitReports.Objects.SummitReportBaseObject).GetTypeInfo().Assembly;
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



                // Generate a Sheet for each relationship.  If uwRelationshipId >0 then onl one sheet is needed.

                string sSQL1 = "";
                if (uwRelationshipId == 0)
                {
                    sSQL1 = @"SET ANSI_WARNINGS OFF; SELECT COUNT(*) AS TabCnt FROM UW.tbl_Relationship WHERE BidPoolId =@p0;";
                }
                else
                {
                    sSQL1 = @"SET ANSI_WARNINGS OFF; SELECT 1 AS TabCnt ;";
                }
                var retTabCnt = await MarsDb.QueryAsDataSetAsync(sSQL1, BidPoolId);
                System.Data.DataTable aResultSet = retTabCnt.Tables[0];
                var iTabCnt = 0;
                foreach (System.Data.DataRow a in aResultSet.Rows)
                { iTabCnt = (int)a["TabCnt"]; }

                for (int x = 2; x < iTabCnt + 1; x++)
                {
                    sheet = workbook.CloneSheet(this.workbook.GetSheetIndex("1"));
                    workbook.SetSheetName(workbook.NumberOfSheets - 1, x.ToString().AsSheetName());
                }
                
                // Return to sheet "1"
                this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex(iSheet.ToString()));

                // Get Dataset for report using ADO;  If uwRelationshipId <> 0 use uwRelationshipId else use BidPoolId

                string sSQL2 = "";
                DataSet retDataSet = null;
                
                if (uwRelationshipId == 0)
                {
                    sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_Loans] WHERE [BidPoolId]=@p0 ORDER BY uwRelationshipId ASC, uwLoanId ASC;";
                    retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, BidPoolId);
                }
                else
                {
                    sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_Loans] WHERE [uwRelationshipId]=@p0 ORDER BY uwRelationshipId ASC, uwLoanId ASC;";
                    retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, uwRelationshipId);
                }
               
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

                    LnCellStyle.WrapText = true;
                    sheet.SetCellValue(iRow + 5, "B", row, "LoanShortName").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "C", row, "LoanDescriptionTxt").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "D", row, "BorrowerTxt").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "E", row, "GuarantorTxt").SetCellStyle(LnCellStyle);
                    LnCellStyle.CellFormat = "mm/dd/yyy";
                    sheet.SetCellValue(iRow + 5, "F", row, "OriginationDate").SetCellStyle(LnCellStyle);
                    if ((row["MaturityDateText"] == System.DBNull.Value) || ((string)row["MaturityDateText"] == ""))
                    {
                        LnCellStyle.CellFormat = "mm/dd/yyy";
                        sheet.SetCellValue(iRow + 5, "G", row, "MaturityDate").SetCellStyle(LnCellStyle);
                    }
                    else 
                    {
                        LnCellStyle.CellFormat = "@";
                        sheet.SetCellValue(iRow + 5, "G", row, "MaturityDatetext").SetCellStyle(LnCellStyle);
                    }
                    LnCellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 5, "H", row, "OriginalUPB").SetCellStyle(LnCellStyle);
                    sheet.SetCellValue(iRow + 5, "I", row, "UPB").SetCellStyle(LnCellStyle);
                    if((row["InterestRateText"] == System.DBNull.Value) || ((string)row["InterestRateText"] == ""))
                    {
                        LnCellStyle.CellFormat = "0.0%";
                        sheet.SetCellValue(iRow + 5, "J", row, "InterestRate").SetCellStyle(LnCellStyle);
                    }
                    else 
                    {
                        LnCellStyle.CellFormat = "@";
                        sheet.SetCellValue(iRow + 5, "J", row, "InterestRateText").SetCellStyle(LnCellStyle);
                    }
                    LnCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 5, "K", row, "SIMValueLoan").SetCellStyle(LnCellStyle);
                    
                    if (iLnCnt == (int)row["LoansCnt"])
                    {
                        sheet.CreateRow(iRow + 7);
                        LnCellStyle.IsBold = true; 
                        sheet.SetCellValue(iRow + 6, "C", "Totals:").SetCellStyle(LnCellStyle);
                        LnCellStyle.CellFormat = "#,##0.00";
                        sheet.SetCellValue(iRow + 6, "H", 0.0).SetCellStyle(LnCellStyle).SetCellFormula(string.Format("SUM(H7:H{0})", (6 + iRow )));
                        sheet.SetCellValue(iRow + 6, "I", 0.0).SetCellStyle(LnCellStyle).SetCellFormula(string.Format("SUM(I7:I{0})", (6 + iRow )));
                        LnCellStyle.CellFormat = "#,###";
                        sheet.SetCellValue(iRow + 6, "K", 0.0).SetCellStyle(LnCellStyle).SetCellFormula(string.Format("SUM(K7:K{0})", (6 + iRow )));
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
