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
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("Sheet1"));
                }
                this.workbook.ClearStyleCache();

                string sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_DeanSheet] WHERE [BidPoolId]=@p0 ORDER BY uwRelationshipId ASC;";
                var dataArr = await MarsDb.Query<UWDeanSheetDTO>(sSQL, BidPoolId);
                var iRow = 1;
                foreach (var data in dataArr)
                {
                    if (iRow == 0)
                    {
                        sheet.SetCellValue(4, "B", data, "BidPool");
                    }
                    else
                    {
                        sheet.CreateRow(iRow + 4);
                        sheet.SetCellValue(iRow + 4, "C", data, "BidSubPoolName");
                        sheet.SetCellValue(iRow + 4, "D", data, "RelationshipName");
                        sheet.SetCellValue(iRow + 4, "F", data, "UW");
                        sheet.SetCellValue(iRow + 4, "G", data, "LoanCount").SetCellFormat("0");
                        sheet.SetCellValue(iRow + 4, "H", data, "UPBSum").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "I", data, "BidAmount").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "J", data, "BidUPB").SetCellFormat("0.00%");
                        sheet.SetCellValue(iRow + 4, "K", data, "DiscountRate").SetCellFormat("0.00%");
                        sheet.SetCellValue(iRow + 4, "L", data, "TrailConC").SetCellFormat("0.00%");
                        sheet.SetCellValue(iRow + 4, "M", data, "ProjConC").SetCellFormat("0.00%");
                        sheet.SetCellValue(iRow + 4, "N", data, "Recovery").SetCellFormat("0.00%");
                        sheet.SetCellValue(iRow + 4, "O", data, "MOIC").SetCellFormat("0.00"); 
                        sheet.SetCellValue(iRow + 4, "P", data, "AppraisalDate");
                        sheet.SetCellValue(iRow + 4, "Q", data, "AppraisalValue").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "R", data, "BusinessAssets").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "S", data, "BankTotal").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "T", data, "BPOValue").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "U", data, "SIMValue").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "V", data, "BidAppr").SetCellFormat("0.00%"); 
                        sheet.SetCellValue(iRow + 4, "W", data, "BidBPO").SetCellFormat("0.00%"); 
                        sheet.SetCellValue(iRow + 4, "X", data, "BidSIMValue").SetCellFormat("0.00%");
                        sheet.SetCellValue(iRow + 4, "Y", data, "PHLast3mth").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "Z", data, "PHLast6mth").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "AA", data, "PHLast9mth").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "AB", data, "PHLast12mth").SetCellFormat("0,000.00");
                        sheet.SetCellValue(iRow + 4, "AC", data, "Recourse");
                        sheet.SetCellValue(iRow + 4, "AD", data, "PrimaryCollateralType");
                        sheet.SetCellValue(iRow + 4, "AE", data, "YearBuilt").SetCellFormat("0000");
                        sheet.SetCellValue(iRow + 4, "AF", data, "REUnit");
                        sheet.SetCellValue(iRow + 4, "AG", data, "REBasis").SetCellFormat("0,000");
                        sheet.SetCellValue(iRow + 4, "AH", data, "PrimaryLocation");
                        sheet.SetCellValue(iRow + 4, "AI", data, "Eyes");
                    }
                    iRow++;
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
