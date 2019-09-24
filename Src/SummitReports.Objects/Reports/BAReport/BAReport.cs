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
using System.Data;

namespace SummitReports.Objects
{
    public class BAReport : SummitReportBaseObject
    {
        public BAReport() : base(@"BAReport\BAReport.xlsx")
        {

        }

        /// <summary>
        /// This will generate a Business Asset Report for a BidPool and Relationship
        /// </summary>
        /// <param name="BidPoolId"></param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> BidPoolGenerateAsync(int BidPoolId)
        {
            return await GenerateAsync(BidPoolId, 0);
        }

        /// <summary>
        /// This will generate a Business Asset Report for a BidPool and Relationship
        /// </summary>
        /// <param name="BidPoolId"></param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> RelationshipGenerateAsync(int uwRElationshipId)
        {
            return await GenerateAsync(0, uwRElationshipId);
        }


        /// <summary>
        /// This will generate a Business Asset Report for a BidPool and Relationship
        /// </summary>
        /// <param name="BidPoolId">If this is 0, then we will assume that we are going to use uwRelationshipId</param>
        /// <param name="uwRelationshipId">If this is zero, then we will assume that we are going top use BidPoolId</param>
        /// <returns>Name of the file generated</returns>
        private async Task<string> GenerateAsync(int BidPoolId, int uwRelationshipId)
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

                // Generate a Sheet for each relationship.  If uwRelationshipId <> 0 then only 1 sheet is needed.

                string  sSQL1 = "";
                if (uwRelationshipId == 0)
                {
                    sSQL1 = @"SET ANSI_WARNINGS OFF; SELECT COUNT(*) AS TabCnt FROM (SELECT DISTINCT r.uwRelationshipId FROM UW.tbl_Relationship AS r INNER JOIN UW.tbl_CollateralNRE AS c ON r.uwRelationshipId = c.uwRelationshipId WHERE r.BidPoolId =@p0) AS a;";
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
                    workbook.SetSheetName(workbook.NumberOfSheets - 1, x.ToString());
                }
                
                // Return to sheet "1"
                this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex(iSheet.ToString()));

                // Get Dataset for report using ADO;  If uwRelationshipId <> 0 use uwRelationshipId else use BidPoolId
                string sSQL2 = "";
                DataSet retDataSet = null;

                if (uwRelationshipId == 0)
                {
                    sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_CollateralNRE] WHERE [BidPoolId]=@p0 ORDER BY uwRelationshipId ASC, uwNRECollateralId ASC;";
                    retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, BidPoolId);
                }
                else
                {
                    sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_CollateralNRE] WHERE [uwRelationshipId]=@p0 ORDER BY uwRelationshipId ASC, uwNRECollateralId ASC;";
                    retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, uwRelationshipId);
                }
                System.Data.DataTable firstResultSet = retDataSet.Tables[0];
                var iRow = 1;
                var iRel = 0;
                var iNRECnt = 1;
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
                        iNRECnt = 1;
                        iRel = (int)row["uwRelationshipId"];
                    }

                    var formatStr = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                    var BACellStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, CellFormat = formatStr, VerticalAlignment = VerticalAlignment.Top, HorizontalAlignment = HorizontalAlignment.Left };

                    
                    sheet.SetCellValue(2, "B", row, "RptHeader");
                    sheet.CreateRow(iRow + 5);
                    sheet.SetCellValue(iRow + 5, "B", row, "NREItemLabel").SetCellStyle(BACellStyle);
                    sheet.SetCellValue(iRow + 5, "C", row, "NREItemComments").SetCellStyle(BACellStyle);
                    BACellStyle.CellFormat = "#,###.00";
                    sheet.SetCellValue(iRow + 5, "D", row, "NREItemBookVal").SetCellStyle(BACellStyle);
                    BACellStyle.CellFormat = "0.0%";
                    sheet.SetCellValue(iRow + 5, "E", row, "NREItemCollPcnt").SetCellStyle(BACellStyle);
                    BACellStyle.CellFormat = "#,###.00";
                    sheet.SetCellValue(iRow + 5, "F", row, "NRESIM").SetCellStyle(BACellStyle);

                    if (iNRECnt == (int)row["CollateralNRECnt"])
                    {
                        //sheet.CreateRow(18 + iRow);
                        //sheet.SetCellValue(18 + iRow, "C", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(C18:C{0})", (18 + iRow - 2)));
                        sheet.CreateRow(iRow + 7);
                        BACellStyle.IsBold = true; 
                        sheet.SetCellValue(iRow + 6, "C", "Totals:").SetCellStyle(BACellStyle);
                        BACellStyle.CellFormat = "#,###.00";
                        sheet.SetCellValue(iRow + 6, "D", 0.0).SetCellStyle(BACellStyle).SetCellFormula(string.Format("SUM(D6:D{0})", (6 + iRow)));
                        sheet.SetCellValue(iRow + 6, "F", 0.0).SetCellStyle(BACellStyle).SetCellFormula(string.Format("SUM(F6:F{0})", (6 + iRow)));

                        BACellStyle.IsBold = false;

                    }

                    iRow++;
                    iNRECnt++;

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
