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

                /* Using ADO.NET Specified */
                string sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT top(16) * FROM [UW].[vw_CollateralRE] WHERE [BidPoolId]=@p0 ORDER BY uwRelationshipId ASC, uwRECollateralId ASC;";
                var retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, BidPoolId);
                System.Data.DataTable firstResultSet = retDataSet.Tables[0];
                var iRow = 0;
                var iRel = 0;
                var iColCnt = 1;
                foreach (System.Data.DataRow row in firstResultSet.Rows)
                {
                    
                    if (iRow ==0)
                    {
                        iRel = (int) row["uwRelationshipId"];
                    }
                    else if (iRel != (int)row["uwRelationshipId"])
                    {
                        iSheet++;
                        this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex(iSheet.ToString()));
                        iRow = 0;
                        iColCnt = 1;
                        iRel = (int)row["uwRelationshipId"];
                    }

                    var formatStr = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                    var RECellStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, CellFormat = formatStr };

                    RECellStyle.CellFormat = "@";
                    sheet.SetCellValue(3, "B", row, "RptHeader");
                    sheet.CreateRow(iRow + 8);
                    sheet.SetCellValue(iRow + 8, "B", row, "CollateralDescriptionTxt").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 8, "C", row, "Size").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "@";
                    sheet.SetCellValue(iRow + 8, "D", row, "SizeMetricDesc").SetCellStyle(RECellStyle); 
                    sheet.SetCellValue(iRow + 8, "E", row, "CollateralFullAddress").SetCellStyle(RECellStyle); 
                    sheet.SetCellValue(iRow + 8, "F", row, "Comments").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "mm/dd/yyy";
                    sheet.SetCellValue(iRow + 8, "G", row, "MRAppraisalDate").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 8, "H", row, "MRAppraisalValue").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 8, "I", row, "MRAppraisalValuetoMetric").SetCellStyle(RECellStyle);
                    RECellStyle.CellFormat = "#,##0.00";
                    sheet.SetCellValue(iRow + 8, "J", row, "BPOValueCRE").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 8, "K", row, "BPOValueCREtoMetric").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 8, "L", row, "SIMValue").SetCellStyle(RECellStyle);
                    sheet.SetCellValue(iRow + 8, "M", row, "SIMValuetoMetric").SetCellStyle(RECellStyle);

                    if (iColCnt == (int)row["CollateralRECnt"])
                    {
                        //sheet.CreateRow(18 + iRow);
                        //sheet.SetCellValue(18 + iRow, "C", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(C18:C{0})", (18 + iRow - 2)));
                        sheet.CreateRow(iRow + 9);
                        RECellStyle.IsBold = true; 
                        sheet.SetCellValue(iRow + 9, "B", "TotalS:").SetCellStyle(RECellStyle);
                        sheet.SetCellValue(iRow + 9, "H", 0.0).SetCellStyle(RECellStyle).SetCellFormula(string.Format("SUM(H8:H{0})", (9 + iRow - 1)));
                        sheet.SetCellValue(iRow + 9, "J", 0.0).SetCellStyle(RECellStyle).SetCellFormula(string.Format("SUM(J8:J{0})", (9 + iRow - 1)));
                        sheet.SetCellValue(iRow + 9, "L", 0.0).SetCellStyle(RECellStyle).SetCellFormula(string.Format("SUM(L8:L{0})", (9 + iRow - 1)));
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
