using System;
using NPOI.SS.UserModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using SummitReports.Infrastructure;

namespace SummitReports.Objects
{
    public class DeanSheetPresentation : SummitReportBaseObject, IBidPoolReport
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

                var assembly = typeof(SummitReports.Objects.SummitReportBaseObject).GetTypeInfo().Assembly;
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



                var formatStr = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                var DSCellStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, CellFormat = formatStr};
                                               
                foreach (System.Data.DataRow row in resultSet.Rows)
                {
                    if (iRow == 1) // Populate Bid Pool Header
                    {
                        sheet.SetCellValue(1, "B", row, "BidPool");
                    }

                    sheet.CreateRow(iRow + 3);
                    DSCellStyle.CellFormat = "@";
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Left;
                    sheet.SetCellValue(iRow + 3, "B", row, "RelationshipName").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "C", row, "BidSubPoolName").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "D", row, "UW").SetCellStyle(DSCellStyle);
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 3, "E", row, "LoanCount").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 3, "F", row, "UPBSum").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "G", row, "BidAmount").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "0.0%";
                    sheet.SetCellValue(iRow + 3, "H", row, "BidUPB").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "I", row, "DiscountRate").SetCellStyle(DSCellStyle);
                    if ((double)row["TrailConC"] == -1)
                    {
                        DSCellStyle.CellFormat = "@";
                        //DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                        sheet.SetCellValue(iRow + 3, "J", "N/A").SetCellStyle(DSCellStyle);
                        DSCellStyle.CellFormat = "0.0%";
                    }
                    else
                    {
                        sheet.SetCellValue(iRow + 3, "J", row, "TrailConC").SetCellStyle(DSCellStyle);
                    }
                    if ((double)row["ProjConC"] == -1)
                    {
                        DSCellStyle.CellFormat = "@";
                        //DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                        sheet.SetCellValue(iRow + 3, "K", "N/A").SetCellStyle(DSCellStyle);
                        DSCellStyle.CellFormat = "0.0%";
                    }
                    else
                    {
                        sheet.SetCellValue(iRow + 3, "K", row, "ProjConC").SetCellStyle(DSCellStyle);
                    }
                    sheet.SetCellValue(iRow + 3, "L", row, "Recovery").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###.00";
                    sheet.SetCellValue(iRow + 3, "M", row, "MOIC").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "mm/dd/yyyy";
                    sheet.SetCellValue(iRow + 3, "N", row, "AppraisalDate").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 3, "O", row, "AppraisalValue").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "P", row, "BusinessAssets").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "Q", row, "BankTotal").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "R", row, "BPOValue").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "S", row, "SIMValue").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "0.0%";
                    if ((double)row["BidAppr"] == -1)
                    {
                        DSCellStyle.CellFormat = "@";
                        //DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                        sheet.SetCellValue(iRow + 3, "T", "N/A").SetCellStyle(DSCellStyle);
                        DSCellStyle.CellFormat = "0.0%";
                    }
                    else
                    {
                        sheet.SetCellValue(iRow + 3, "T", row, "BidAppr").SetCellStyle(DSCellStyle);
                    }
                    if ((double)row["BidBPO"] == -1)
                    {
                        DSCellStyle.CellFormat = "@";
                        //DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                        sheet.SetCellValue(iRow + 3, "U", "N/A").SetCellStyle(DSCellStyle);
                        DSCellStyle.CellFormat = "0.0%";
                    }
                    else
                    {
                        sheet.SetCellValue(iRow + 3, "U", row, "BidBPO").SetCellStyle(DSCellStyle);
                    }
                    if ((double)row["BidSIMValue"] == -1)
                    {
                        DSCellStyle.CellFormat = "@";
                        //DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                        sheet.SetCellValue(iRow + 3, "V", "N/A").SetCellStyle(DSCellStyle);
                        DSCellStyle.CellFormat = "0.0%";
                    }
                    else
                    {
                        sheet.SetCellValue(iRow + 3, "V", row, "BidSIMValue").SetCellStyle(DSCellStyle);
                    }
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 3, "W", row, "PHLast3mth").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "X", row, "PHLast6mth").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "Y", row, "PHLast9mth").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 3, "Z", row, "PHLast12mth").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "@";
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Center;
                    sheet.SetCellValue(iRow + 3, "AA", row, "Recourse").SetCellStyle(DSCellStyle);
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Left;
                    sheet.SetCellValue(iRow + 3, "AB", row, "PrimaryCollateralType").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#";
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                    sheet.SetCellValue(iRow + 3, "AC", row, "YearBuilt").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "@";
                    sheet.SetCellValue(iRow + 3, "AD", row, "REUnit").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 3, "AE", row, "REBasis").SetCellStyle(DSCellStyle); 
                    DSCellStyle.CellFormat = "@";
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Left;
                    sheet.SetCellValue(iRow + 3, "AF", row, "PrimaryLocation").SetCellStyle(DSCellStyle);
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Center;
                    sheet.SetCellValue(iRow + 3, "AG", row, "Eyes").SetCellStyle(DSCellStyle);

                    iRow++;
                }

                resultSet = retDataSet.Tables[1];
                foreach (System.Data.DataRow row in resultSet.Rows)
                {
                    sheet.CreateRow(iRow + 4);
                    DSCellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                    DSCellStyle.IsBold = true;
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 4, "C", "Totals:").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "D", "").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 4, "E", row, "LoanCount").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 4, "F", row, "UPBSum").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "G", row, "BidAmount").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "0.0%";
                    sheet.SetCellValue(iRow + 4, "H", row, "BidUPB").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "I", row, "DiscountRate").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "J", row, "TrailConC").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "K", row, "ProjConC").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "L", row, "Recovery").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###.00";
                    sheet.SetCellValue(iRow + 4, "M", row, "MOIC").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "N", "").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 4, "O", row, "AppraisalValue").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "P", row, "BusinessAssets").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "Q", row, "BankTotal").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "R", row, "BPOValue").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "S", row, "SIMValue").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "0.0%";
                    sheet.SetCellValue(iRow + 4, "T", row, "BidAppr").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "U", row, "BidBPO").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "V", row, "BidSIMValue").SetCellStyle(DSCellStyle);
                    DSCellStyle.CellFormat = "#,###";
                    sheet.SetCellValue(iRow + 4, "W", row, "PHLast3mth").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "X", row, "PHLast6mth").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "Y", row, "PHLast9mth").SetCellStyle(DSCellStyle);
                    sheet.SetCellValue(iRow + 4, "Z", row, "PHLast12mth").SetCellStyle(DSCellStyle);
                    DSCellStyle.IsBold = false;
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
