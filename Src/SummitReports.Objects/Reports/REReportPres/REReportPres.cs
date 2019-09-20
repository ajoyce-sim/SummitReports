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
    public class REReportPres : SummitReportBaseObject
    {
        public REReportPres() : base(@"REReportPres\REReportPres.xlsx")
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

                // Generate a Sheet for each relationship

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
                string sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_CollateralRE] WHERE [BidPoolId]=@p0 ORDER BY uwRelationshipId ASC, uwRECollateralId ASC;";
                var retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, BidPoolId);
                System.Data.DataTable firstResultSet = retDataSet.Tables[0];
                var iRow = 1;
                var iRel = 0;
                var iColCnt = 1;
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
                        iColCnt = 1;
                        iRel = (int)row["uwRelationshipId"];
                    }

                    var formatStr = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                    var RECellStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, CellFormat = formatStr, VerticalAlignment = VerticalAlignment.Top, HorizontalAlignment = HorizontalAlignment.Left };

                    RECellStyle.CellFormat = "@";
                    sheet.SetCellValue(2, "B", row, "RptHeader");
                    sheet.CreateRow(iRow + 6);
                    RECellStyle.WrapText = true;
                    sheet.SetCellValue(iRow + 6, "B", row, "CollateralDescriptionTxt").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 6, "C", row, "Size").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "@";
                    sheet.SetCellValue(iRow + 6, "D", row, "SizeMetricDesc").SetCellStyle(RECellStyle); 
                    sheet.SetCellValue(iRow + 6, "E", row, "CollateralFullAddress").SetCellStyle(RECellStyle); 
                    sheet.SetCellValue(iRow + 6, "F", row, "Comments").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "mm/dd/yyy";
                    RECellStyle.WrapText = false;
                    sheet.SetCellValue(iRow + 6, "G", row, "MRAppraisalDate").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 6, "H", row, "MRAppraisalValue").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 6, "I", row, "MRAppraisalValuetoMetric").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 6, "J", row, "BPOValueCRE").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 6, "K", row, "BPOValueCREtoMetric").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 6, "L", row, "SIMValue").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 6, "M", row, "SIMValuetoMetric").SetCellStyle(RECellStyle);
                    

                    if (iColCnt == (int)row["CollateralRECnt"])
                    {
                        //sheet.CreateRow(18 + iRow);
                        //sheet.SetCellValue(18 + iRow, "C", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(C18:C{0})", (18 + iRow - 2)));
                        sheet.CreateRow(iRow + 7);
                        RECellStyle.IsBold = true; 
                        sheet.SetCellValue(iRow + 7, "B", "TotalS:").SetCellStyle(RECellStyle);
                        sheet.SetCellValue(iRow + 7, "H", 0.0).SetCellStyle(RECellStyle).SetCellFormula(string.Format("SUM(H8:H{0})", (7 + iRow )));
                        sheet.SetCellValue(iRow + 7, "J", 0.0).SetCellStyle(RECellStyle).SetCellFormula(string.Format("SUM(J8:J{0})", (7 + iRow )));
                        sheet.SetCellValue(iRow + 7, "L", 0.0).SetCellStyle(RECellStyle).SetCellFormula(string.Format("SUM(L8:L{0})", (7 + iRow )));
                        RECellStyle.IsBold = false;
                    }

                    iRow++;
                    iColCnt++;

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
