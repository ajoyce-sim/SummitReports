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

namespace SummitReports.Objects
{
    public class UWRelationshipCashFlow : SummitReportBaseObject
    {
        public UWRelationshipCashFlow() : base(@"UWRelationshipCashFlowReport\UW-RCF-Reports.xlsx")
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="uwRelationshipId"></param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int uwRelationshipId)
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
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("Relationship Cash Flow"));
                }
                this.workbook.ClearStyleCache();


                var svc = new UWDataService();
                var reldata = await svc.FetchUWRelationshipData(uwRelationshipId);
                var relCFdata = await svc.FetchUWRelationshipCashFlowsData(uwRelationshipId);
                foreach (var uwRelItem in reldata)
                {
                    sheet.SetCellValue(2, "B", uwRelItem, "uwRelationshipId");
                    sheet.SetCellValue(2, "E", uwRelItem, "Underwriter");
                    sheet.SetCellValue(2, "G", uwRelItem, "PrimaryCollateralType");

                    sheet.SetCellValue(3, "B", uwRelItem, "RelationshipName");
                    sheet.SetCellValue(3, "E", uwRelItem, "SiteVisitFlag");
                    sheet.SetCellValue(3, "G", uwRelItem, "RecourseFlag");
                    sheet.SetCellValue(3, "I", uwRelItem, "TotalSIM");
                    sheet.SetCellValue(3, "K", uwRelItem, "TotalApprBook");
                    sheet.SetCellValue(3, "M", uwRelItem, "UPBSum");
                    sheet.SetCellValue(3, "N", uwRelItem, "BidAmount");
                    sheet.SetCellValue(3, "O", uwRelItem, "BidUPB");
                    sheet.SetCellValue(3, "P", uwRelItem, "BidSIM");

                    sheet.SetCellValue(5, "B", uwRelItem, "BidPoolName");
                    sheet.SetCellValue(5, "G", uwRelItem, "YearBuilt");

                    sheet.SetCellValue(6, "B", uwRelItem, "BidSubPoolName");
                    sheet.SetCellValue(6, "G", uwRelItem, "BPOValueCRE");
                    sheet.SetCellValue(6, "I", uwRelItem, "SIMValue");
                    sheet.SetCellValue(6, "J", uwRelItem, "AppraisalDate");
                    sheet.SetCellValue(6, "K", uwRelItem, "AppraisalValue");
                    sheet.SetCellValue(6, "M", uwRelItem, "DiscountRate");
                    sheet.SetCellValue(6, "N", uwRelItem, "Recovery");
                    sheet.SetCellValue(6, "O", uwRelItem, "MOIC");
                    sheet.SetCellValue(6, "P", uwRelItem, "WAL");

                    sheet.SetCellValue(7, "B", uwRelItem, "PrimaryCity");
                    sheet.SetCellValue(7, "D", uwRelItem, "PrimaryCounty");
                    sheet.SetCellValue(7, "F", uwRelItem, "PrimaryState");

                    sheet.SetCellValue(8, "B", uwRelItem, "CollateralDescText");
                    sheet.SetCellValue(9, "I", 0.0);
                    sheet.SetCellValue(9, "M", uwRelItem, "CashOnCash");
                    sheet.SetCellValue(9, "N", uwRelItem, "GrossCashFlow");
                    sheet.SetCellValue(9, "O", uwRelItem, "NetCashFlow");
                    sheet.SetCellValue(9, "P", uwRelItem, "LegalGross");

                    sheet.SetCellValue(10, "B", uwRelItem, "PrimaryLienPoisition");
                    sheet.SetCellValue(10, "E", uwRelItem, "ExitTypeDesc");

                    sheet.SetCellValue(11, "B", uwRelItem, "CurrentStatus");
                    sheet.SetCellValue(11, "E", uwRelItem, "ExitStrategyText");

                    sheet.SetCellValue(12, "B", uwRelItem, "ProFormaStatus");
                    sheet.SetCellValue(12, "M", uwRelItem, "PHLast3mth");
                    sheet.SetCellValue(12, "N", uwRelItem, "PHLast6mth");
                    sheet.SetCellValue(12, "O", uwRelItem, "PHLast9mth");
                    sheet.SetCellValue(12, "P", uwRelItem, "PHLast12mth");

                    sheet.SetCellValue(13, "B", uwRelItem, "PerformingRate");
                    sheet.SetCellValue(13, "I", uwRelItem, "Size");
                    sheet.SetCellValue(13, "J", uwRelItem, "SizeMetricDesc");
                    sheet.SetCellValue(13, "K", uwRelItem, "BidMetric");
                    sheet.SetCellValue(13, "L", uwRelItem, "SIMMetric");

                    sheet.SetCellValue(17, "E", uwRelItem, "MiscIncome3Label");
                    sheet.SetCellValue(17, "F", uwRelItem, "MiscIncome4Label");
                    sheet.SetCellValue(17, "G", uwRelItem, "MiscIncome5Label");
                    sheet.SetCellValue(17, "H", uwRelItem, "MiscIncome6Label");

                    var formatStr = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                    var cashFlowCellStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, CellFormat = formatStr };

                    var iRow = 0;
                    foreach (var cashFlowItem in relCFdata)
                    {
                        sheet.CreateRow(18 + iRow);
                        
                        sheet.SetCellValue(18 + iRow, "A", cashFlowItem, "CashFlowDate").SetCellFormat("mmm-yy");
                        sheet.SetCellValue(18 + iRow, "B", (double)(iRow + 1)).SetCellFormat("0"); ;
                        sheet.SetCellValue(18 + iRow, "C", cashFlowItem, "Principal").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "D", cashFlowItem, "Interest").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "E", cashFlowItem, "MiscIncome3").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "F", cashFlowItem, "MiscIncome4").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "G", cashFlowItem, "MiscIncome5").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "H", cashFlowItem, "MiscIncome6").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "I", (double)(iRow + 1)).SetCellFormat("0"); ;
                        sheet.SetCellValue(18 + iRow, "J", cashFlowItem, "BackTaxes").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "K", cashFlowItem, "Legal").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "L", cashFlowItem, "Travel").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "M", cashFlowItem, "BrokerFee").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "N", cashFlowItem, "REOTax").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "O", cashFlowItem, "REOins").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "P", cashFlowItem, "CapEx").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "Q", cashFlowItem, "TiLc").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "R", cashFlowItem, "Environ").SetCellStyle(cashFlowCellStyle);
                        sheet.SetCellValue(18 + iRow, "S", cashFlowItem, "Misc").SetCellStyle(cashFlowCellStyle);
                        iRow++;
                    }
                    iRow++;

                    sheet.CreateRow(18 + iRow);
                    sheet.SetCellValue(18 + iRow, "C", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(C18:C{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "D", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(D18:D{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "E", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(E18:E{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "F", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(F18:F{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "G", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(G18:G{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "H", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(H18:H{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "J", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(I18:I{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "K", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(J18:J{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "L", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(K18:K{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "M", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(L18:L{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "N", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(M18:M{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "O", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(N18:N{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "P", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(O18:O{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "Q", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(P18:P{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "R", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(Q18:Q{0})", (18 + iRow - 2)));
                    sheet.SetCellValue(18 + iRow, "S", 0.0).SetCellFormat(formatStr).SetCellFormula(string.Format("SUM(R18:R{0})", (18 + iRow - 2)));

                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("Notes"));

                    sheet.SetCellValue(2, "A", uwRelItem, "AssetNotes");
                    sheet.SetCellValue(23, "A", uwRelItem, "CollateralValuationNotes");
                    sheet.SetCellValue(45, "A", uwRelItem, "TitleUCCNotes");
                    sheet.SetCellValue(57, "A", uwRelItem, "EnvironmentalNotes");

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
