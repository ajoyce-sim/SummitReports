using System;
using NPOI.SS.UserModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using System.Data;
using SummitReport.Infrastructure;

namespace SummitReports.Objects
{
    public class UWRelationshipCashFlow : SummitReportBaseObject, IBidPoolRelationshipReport
    {
        public UWRelationshipCashFlow() : base(@"UWRelationshipCashFlowReport\UW-RCF-Reports.xlsx")
        {

        }
        
        private bool GenerateSheetForRelationship(ISheet sheet, DataRow uwRelItem, DataTable _relCFdata)
        {
            try
            {
                sheet.SetCellValue(1, "B", uwRelItem, "uwRelationshipId");
                sheet.SetCellValue(1, "E", uwRelItem, "Underwriter");
                sheet.SetCellValue(1, "G", uwRelItem, "PrimaryCollateralType");

                sheet.SetCellValue(2, "B", uwRelItem, "RelationshipName");
                sheet.SetCellValue(2, "E", uwRelItem, "SiteVisitFlag");
                sheet.SetCellValue(2, "G", uwRelItem, "RecourseFlag");
                sheet.SetCellValue(2, "I", uwRelItem, "TotalSIM");
                sheet.SetCellValue(2, "K", uwRelItem, "TotalApprBook");
                sheet.SetCellValue(2, "M", uwRelItem, "UPBSum");
                sheet.SetCellValue(2, "N", uwRelItem, "BidAmount");
                sheet.SetCellValue(2, "O", uwRelItem, "BidUPB");
                sheet.SetCellValue(2, "P", uwRelItem, "BidSIM");

                sheet.SetCellValue(4, "B", uwRelItem, "BidPoolName");
                sheet.SetCellValue(4, "G", uwRelItem, "YearBuilt");

                sheet.SetCellValue(5, "B", uwRelItem, "BidSubPoolName");
                sheet.SetCellValue(5, "G", uwRelItem, "BPOValueCRE");
                sheet.SetCellValue(5, "I", uwRelItem, "SIMValue");
                sheet.SetCellValue(5, "J", uwRelItem, "AppraisalDate");
                sheet.SetCellValue(5, "K", uwRelItem, "AppraisalValue");
                sheet.SetCellValue(5, "M", uwRelItem, "DiscountRate");
                sheet.SetCellValue(5, "N", uwRelItem, "Recovery");
                sheet.SetCellValue(5, "O", uwRelItem, "MOIC");
                sheet.SetCellValue(5, "P", uwRelItem, "WAL");

                sheet.SetCellValue(6, "B", uwRelItem, "PrimaryCity");
                sheet.SetCellValue(6, "D", uwRelItem, "PrimaryCounty");
                sheet.SetCellValue(6, "F", uwRelItem, "PrimaryState");

                sheet.SetCellValue(7, "B", uwRelItem, "CollateralDescText");
                sheet.SetCellValue(8, "I", 0.0);
                sheet.SetCellValue(8, "M", uwRelItem, "CashOnCash");
                sheet.SetCellValue(8, "N", uwRelItem, "GrossCashFlow");
                sheet.SetCellValue(8, "O", uwRelItem, "NetCashFlow");
                sheet.SetCellValue(8, "P", uwRelItem, "LegalGross");

                sheet.SetCellValue(9, "B", uwRelItem, "PrimaryLienPoisition");
                sheet.SetCellValue(9, "E", uwRelItem, "ExitTypeDesc");

                sheet.SetCellValue(10, "B", uwRelItem, "CurrentStatus");
                sheet.SetCellValue(10, "E", uwRelItem, "ExitStrategyText");

                sheet.SetCellValue(11, "B", uwRelItem, "ProFormaStatus");
                sheet.SetCellValue(11, "M", uwRelItem, "PHLast3mth");
                sheet.SetCellValue(11, "N", uwRelItem, "PHLast6mth");
                sheet.SetCellValue(11, "O", uwRelItem, "PHLast9mth");
                sheet.SetCellValue(11, "P", uwRelItem, "PHLast12mth");

                sheet.SetCellValue(12, "B", uwRelItem, "PerformingRate");
                sheet.SetCellValue(12, "I", uwRelItem, "Size");
                sheet.SetCellValue(12, "J", uwRelItem, "SizeMetricDesc");
                sheet.SetCellValue(12, "K", uwRelItem, "BidMetric");
                sheet.SetCellValue(12, "L", uwRelItem, "SIMMetric");

                sheet.SetCellValue(16, "E", uwRelItem, "MiscIncome3Label");
                sheet.SetCellValue(16, "F", uwRelItem, "MiscIncome4Label");
                sheet.SetCellValue(16, "G", uwRelItem, "MiscIncome5Label");
                sheet.SetCellValue(16, "H", uwRelItem, "MiscIncome6Label");

                var formatStr = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                var cashFlowCellStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, CellFormat = formatStr };

                var iRow = 0;
                var relCFdata = _relCFdata.Select(string.Format("uwRelationshipId={0}", int.Parse(uwRelItem["uwRelationshipId"].ToString())));
                foreach (var cashFlowItem in relCFdata)
                {
                    sheet.CreateRow(17 + iRow);

                    sheet.SetCellValue(17 + iRow, "A", cashFlowItem, "CashFlowDate").SetCellFormat("mmm-yy");
                    sheet.SetCellValue(17 + iRow, "B", (double)(iRow + 1)).SetCellFormat("0"); ;
                    sheet.SetCellValue(17 + iRow, "C", cashFlowItem, "Principal").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "D", cashFlowItem, "Interest").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "E", cashFlowItem, "MiscIncome3").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "F", cashFlowItem, "MiscIncome4").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "G", cashFlowItem, "MiscIncome5").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "H", cashFlowItem, "MiscIncome6").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "I", (double)(iRow + 1)).SetCellFormat("0"); ;
                    sheet.SetCellValue(17 + iRow, "J", cashFlowItem, "BackTaxes").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "K", cashFlowItem, "Legal").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "L", cashFlowItem, "Travel").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "M", cashFlowItem, "BrokerFee").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "N", cashFlowItem, "REOTax").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "O", cashFlowItem, "REOins").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "P", cashFlowItem, "CapEx").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "Q", cashFlowItem, "TiLc").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "R", cashFlowItem, "Environ").SetCellStyle(cashFlowCellStyle);
                    sheet.SetCellValue(17 + iRow, "S", cashFlowItem, "Misc").SetCellStyle(cashFlowCellStyle);
                    iRow++;
                }
                iRow++;

                sheet.CreateRow(17 + iRow);
                sheet.SetCellValue(17 + iRow, "C", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(C18:C{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "D", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(D18:D{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "E", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(E18:E{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "F", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(F18:F{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "G", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(G18:G{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "H", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(H18:H{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "J", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(I18:I{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "K", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(J18:J{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "L", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(K18:K{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "M", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(L18:L{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "N", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(M18:M{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "O", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(N18:N{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "P", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(O18:O{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "Q", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(P18:P{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "R", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(Q18:Q{0})", (18 + iRow - 2)));
                sheet.SetCellValue(17 + iRow, "S", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(R18:R{0})", (18 + iRow - 2)));

                iRow = iRow + 17 + 3; //line up past the model last row and the last row of all the generated rows

                var titleStyle = new XSSFNPoiStyle() {  IsBold = true, FontHeightInPoints = 30  };
                var notesStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, HorizontalAlignment= HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Top, WrapText = true };

                sheet.SetCellValue(iRow, "A", "Asset Notes").SetCellStyle(titleStyle);
                iRow++;
                this.sheet.MergeCellsRange(iRow, iRow + 21, "A", "Z").SetRangeStyle(this.sheet, notesStyle);
                sheet.SetCellValue(iRow, "A", uwRelItem, "AssetNotes").SetCellStyle(notesStyle);
                iRow = iRow + 22;

                sheet.SetCellValue(iRow, "A", "Collateral Valuation Notes").SetCellStyle(titleStyle);
                iRow++;
                this.sheet.MergeCellsRange(iRow, iRow + 21, "A", "Z").SetRangeStyle(this.sheet, notesStyle);
                sheet.SetCellValue(iRow, "A", uwRelItem, "CollateralValuationNotes").SetCellStyle(notesStyle);
                iRow = iRow + 22;

                sheet.SetCellValue(iRow, "A", "Title UCC Notes").SetCellStyle(titleStyle);
                iRow++;
                this.sheet.MergeCellsRange(iRow, iRow + 21, "A", "Z").SetRangeStyle(this.sheet, notesStyle);
                sheet.SetCellValue(iRow, "A", uwRelItem, "TitleUCCNotes").SetCellStyle(notesStyle);
                iRow = iRow + 22;

                sheet.SetCellValue(iRow, "A", "Environmental Notes").SetCellStyle(titleStyle);
                iRow++;
                this.sheet.MergeCellsRange(iRow, iRow + 21, "A", "Z").SetRangeStyle(this.sheet, notesStyle);
                sheet.SetCellValue(iRow, "A", uwRelItem, "EnvironmentalNotes").SetCellStyle(notesStyle);
                iRow = iRow + 22;
            }
            catch (Exception)
            {
                throw;
            }
            return true;
        }

        /// <summary>
        /// Generate Cash Flow Reports for all the relationships in a Bid Pool
        /// </summary>
        /// <param name="BidPoolId">The Bid Pool with the relationships cash flow report to generate</param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> BidPoolGenerateAsync(int BidPoolId)
        {
            return await GenerateAsync(BidPoolId, 0);
        }
        /// <summary>
        /// Generate Cash Flow Reports for all the relationships in a Bid Pool
        /// </summary>
        /// <returns>Name of the file generated</returns>
        /// <returns></returns>
        public async Task<string> RelationshipGenerateAsync(int uwRelationshipId)
        {
            return await GenerateAsync(0, uwRelationshipId);
        }

        /// <summary>
        /// Generate Cashflow report By either BidPool or Relationship Id
        /// </summary>
        /// <param name="BidPoolId">If this is > 0 AND uwRelationshipId = 0,  then this app will generate cash flow sheets for every relationship in a bid pool</param>
        /// <param name="uwRelationshipId">If this is > 0 AND BidPoolId = 0,  then this app will generate cash flow sheets for a single bit pool</param>
        /// <returns>Name of the file generated</returns>
        private async Task<string> GenerateAsync(int BidPoolId, int uwRelationshipId)
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
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("MODEL"));
                }
                this.workbook.ClearStyleCache();

                DataSet retDataSet = null;
                if ((BidPoolId==0) && (uwRelationshipId>0))
                {
                    string sSQL = @"SET ANSI_WARNINGS OFF; 
SELECT * FROM [UW].[vw_Relationship] WHERE uwRelationshipId = @p0;
SELECT cf.* FROM [UW].[vw_RelationshipCashFlow] cf 
WHERE uwRelationshipId = @p0 ORDER BY CashFlowDate;";
                    retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL, uwRelationshipId);

                }
                else if ((BidPoolId > 0) && (uwRelationshipId == 0))
                {
                    string sSQL = @"SET ANSI_WARNINGS OFF; 
SELECT * FROM [UW].[vw_Relationship] WHERE BidPoolId = @p0 ORDER BY RelationshipName;
SELECT cf.* FROM [UW].[vw_RelationshipCashFlow] cf 
INNER JOIN [UW].[vw_Relationship] r
ON cf.uwRelationshipId = r.uwRelationshipId
WHERE r.BidPoolId = @p0 ORDER BY r.RelationshipName, CashFlowDate;";
                    retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL, BidPoolId);
                } else
                {
                    throw new Exception(string.Format("Invalid call to GenerateAsync is invalid.  We need one or the other to be greater than zero.  BidPoolId={0} uwRelationshipId={1}", BidPoolId, uwRelationshipId));
                }
                var reldata = retDataSet.Tables[0];
                var relCFDataTable = retDataSet.Tables[1];

                foreach (DataRow uwRelItem in reldata.Rows)
                {
                    sheet = workbook.CloneSheet(this.workbook.GetSheetIndex("MODEL"));
                    workbook.SetSheetName(workbook.NumberOfSheets - 1, uwRelItem["RelationshipName"].ToString().AsSheetName());

                    GenerateSheetForRelationship(this.sheet, uwRelItem, relCFDataTable);
                }

                workbook.RemoveSheetAt(this.workbook.GetSheetIndex("MODEL"));
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
