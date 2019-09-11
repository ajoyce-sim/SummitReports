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
using SummitReports.Objects.Models;

namespace SummitReports.Objects
{
    public class DeanSheetPresentation : SummitReportBaseObject
    {
        public DeanSheetPresentation() : base(@"DeanSheetPresentation\DeanSheet-Presentation.xlsx")
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
                using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
                {
                    this.workbook = new XSSFWorkbook(file);
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("DS"));
                }
                this.workbook.ClearStyleCache();

                string sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_DeanSheet] WHERE [BidPoolId]=@p0 ORDER BY uwRelationshipId ASC;SELECT * FROM [UW].[vw_DeanSheet_Totals] WHERE [BidPoolId]=@p0";
                var retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL, BidPoolId);
                System.Data.DataTable resultSet = retDataSet.Tables[0];
                var iRow = 1;

                // create bordered cell style - cell style curr
                XSSFCellStyle DSCellStyleCurr = (XSSFCellStyle)workbook.CreateCellStyle();
                DSCellStyleCurr.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleCurr.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleCurr.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleCurr.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleCurr.BottomBorderColor = IndexedColors.Black.Index;
                DSCellStyleCurr.TopBorderColor = IndexedColors.Black.Index;
                DSCellStyleCurr.LeftBorderColor = IndexedColors.Black.Index;
                DSCellStyleCurr.RightBorderColor = IndexedColors.Black.Index;
                DSCellStyleCurr.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.BOTTOM, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleCurr.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.LEFT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleCurr.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.TOP, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleCurr.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.RIGHT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                var formatStr = "#,##0.00";
                DSCellStyleCurr.DataFormat = workbook.CreateDataFormat().GetFormat(formatStr);

                // create bordered cell style - cell style pcnt
                XSSFCellStyle DSCellStylePcnt = (XSSFCellStyle)workbook.CreateCellStyle();
                DSCellStylePcnt.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStylePcnt.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStylePcnt.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStylePcnt.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStylePcnt.BottomBorderColor = IndexedColors.Black.Index;
                DSCellStylePcnt.TopBorderColor = IndexedColors.Black.Index;
                DSCellStylePcnt.LeftBorderColor = IndexedColors.Black.Index;
                DSCellStylePcnt.RightBorderColor = IndexedColors.Black.Index;
                DSCellStylePcnt.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.BOTTOM, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStylePcnt.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.LEFT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStylePcnt.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.TOP, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStylePcnt.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.RIGHT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                formatStr = "0.00%";
                DSCellStylePcnt.DataFormat = workbook.CreateDataFormat().GetFormat(formatStr);

                // create bordered cell style - cell style int
                XSSFCellStyle DSCellStyleInt1 = (XSSFCellStyle)workbook.CreateCellStyle();
                DSCellStyleInt1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt1.BottomBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt1.TopBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt1.LeftBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt1.RightBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt1.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.BOTTOM, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleInt1.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.LEFT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleInt1.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.TOP, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleInt1.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.RIGHT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                formatStr = "0";
                DSCellStyleInt1.DataFormat = workbook.CreateDataFormat().GetFormat(formatStr);

                // create bordered cell style - cell style int (,)
                XSSFCellStyle DSCellStyleInt2 = (XSSFCellStyle)workbook.CreateCellStyle();
                DSCellStyleInt2.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt2.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt2.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                DSCellStyleInt2.BottomBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt2.TopBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt2.LeftBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt2.RightBorderColor = IndexedColors.Black.Index;
                DSCellStyleInt2.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.BOTTOM, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleInt2.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.LEFT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleInt2.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.TOP, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                DSCellStyleInt2.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.RIGHT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                formatStr = "#,##0";
                DSCellStyleInt2.DataFormat = workbook.CreateDataFormat().GetFormat(formatStr);

                // create bordered cell style - cell style int date
                XSSFCellStyle cashFlowCellStyleDate = (XSSFCellStyle)workbook.CreateCellStyle();
                cashFlowCellStyleDate.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                cashFlowCellStyleDate.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                cashFlowCellStyleDate.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                cashFlowCellStyleDate.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                cashFlowCellStyleDate.BottomBorderColor = IndexedColors.Black.Index;
                cashFlowCellStyleDate.TopBorderColor = IndexedColors.Black.Index;
                cashFlowCellStyleDate.LeftBorderColor = IndexedColors.Black.Index;
                cashFlowCellStyleDate.RightBorderColor = IndexedColors.Black.Index;
                cashFlowCellStyleDate.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.BOTTOM, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                cashFlowCellStyleDate.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.LEFT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                cashFlowCellStyleDate.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.TOP, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                cashFlowCellStyleDate.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.RIGHT, new NPOI.XSSF.UserModel.XSSFColor(System.Drawing.Color.Black));
                formatStr = "mm/dd/yyyy";
                cashFlowCellStyleDate.DataFormat = workbook.CreateDataFormat().GetFormat(formatStr);
                cashFlowCellStyleDate.DataFormat = workbook.CreateDataFormat().GetFormat(formatStr);
                
                foreach (System.Data.DataRow row in resultSet.Rows)
                {
                    if (iRow == 1) // Populate Bid Pool Header
                    {
                        sheet.SetCellValue(2, "B", row, "BidPool");
                    }

                    sheet.CreateRow(iRow + 4);
                    sheet.SetCellValue(iRow + 4, "B", row, "RelationshipName").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "C", row, "BidSubPoolName").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "D", row, "UW").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "E", row, "LoanCount").SetCellStyle(DSCellStyleInt1);
                    sheet.SetCellValue(iRow + 4, "F", row, "UPBSum").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "G", row, "BidAmount").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "H", row, "BidUPB").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "I", row, "DiscountRate").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "J", row, "TrailConC").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "K", row, "ProjConC").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "L", row, "Recovery").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "M", row, "MOIC").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "N", row, "AppraisalDate").SetCellStyle(cashFlowCellStyleDate);
                    sheet.SetCellValue(iRow + 4, "O", row, "AppraisalValue").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "P", row, "BusinessAssets").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "Q", row, "BankTotal").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "R", row, "BPOValue").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "S", row, "SIMValue").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "T", row, "BidAppr").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "U", row, "BidBPO").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "V", row, "BidSIMValue").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "W", row, "PHLast3mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "X", row, "PHLast6mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "Y", row, "PHLast9mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "Z", row, "PHLast12mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AA", row, "Recourse").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AB", row, "PrimaryCollateralType").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AC", row, "YearBuilt").SetCellStyle(DSCellStyleInt1);
                    sheet.SetCellValue(iRow + 4, "AD", row, "REUnit").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AE", row, "REBasis").SetCellStyle(DSCellStyleInt2);
                    sheet.SetCellValue(iRow + 4, "AF", row, "PrimaryLocation").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AG", row, "Eyes").SetCellStyle(DSCellStyleCurr);

                    iRow++;
                }

                resultSet = retDataSet.Tables[1];
                foreach (System.Data.DataRow row in resultSet.Rows)
                {
                    sheet.CreateRow(iRow + 4);
                    sheet.SetCellValue(iRow + 4, "B", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "C", "Totals:").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "D", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "E", row, "LoanCount").SetCellStyle(DSCellStyleInt1);
                    sheet.SetCellValue(iRow + 4, "E", row, "LoanCount").SetCellStyle(DSCellStyleInt1);
                    sheet.SetCellValue(iRow + 4, "F", row, "UPBSum").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "G", row, "BidAmount").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "H", row, "BidUPB").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "I", row, "DiscountRate").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "J", row, "TrailConC").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "K", row, "ProjConC").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "L", row, "Recovery").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "M", row, "MOIC").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "N", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "O", row, "AppraisalValue").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "P", row, "BusinessAssets").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "Q", row, "BankTotal").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "R", row, "BPOValue").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "S", row, "SIMValue").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "T", row, "BidAppr").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "U", row, "BidBPO").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "V", row, "BidSIMValue").SetCellStyle(DSCellStylePcnt);
                    sheet.SetCellValue(iRow + 4, "W", row, "PHLast3mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "X", row, "PHLast6mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "Y", row, "PHLast9mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "Z", row, "PHLast12mth").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AA", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AB", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AC", "").SetCellStyle(DSCellStyleInt1);
                    sheet.SetCellValue(iRow + 4, "AD", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AE", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AF", "").SetCellStyle(DSCellStyleCurr);
                    sheet.SetCellValue(iRow + 4, "AG", "").SetCellStyle(DSCellStyleCurr);
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
