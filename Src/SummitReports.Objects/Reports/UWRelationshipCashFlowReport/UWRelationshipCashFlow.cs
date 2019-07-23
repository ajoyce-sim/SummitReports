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

namespace SummitReports.Objects
{
    public class UWRelationshipCashFlow
    {
        public UWRelationshipCashFlow()
        {
            string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            this.TemplateFileName = assemblyPath + @"\Reports\UWRelationshipCashFlowReport\UW-RCF-Reports.xlsx";
        }
        private XSSFWorkbook workbook = new XSSFWorkbook();
        private ISheet sheet;

        private int rowIndex = 0;
        public string TemplateFileName = "";
        public string GeneratedFileName = ""; 
        private void SetupWorksheet()
        {
            this.GeneratedFileName = System.IO.Path.GetTempPath() + "UW-RCF-Reports-" + Guid.NewGuid().ToString() + ".xlsx";
            File.Copy(this.TemplateFileName, this.GeneratedFileName);
            using (FileStream file = new FileStream(this.GeneratedFileName, FileMode.Open, FileAccess.Read))
            {
                this.workbook = new XSSFWorkbook(file);
                this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("Relationship Cash Flow"));
            }

        }
        public bool SaveToFile(string FileName)
        {
            using (var file2 = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite))
            {
                workbook.Write(file2);
                file2.Close();
            }
            return true;
        }
        public void Clear()
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
                SetupWorksheet();
                var svc = new UWDataService();
                var reldata = await svc.FetchUWRelationshipData(13);
                var relCFdata = await svc.FetchUWRelationshipCashFlowsData(13);
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

                    sheet.SetCellValue(9, "B", uwRelItem, "CollateralDescText");
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


                    sheet.SetCellValue(71, "A", uwRelItem, "AssetNotes");
                    sheet.SetCellValue(92, "A", uwRelItem, "CollateralValuationNotes");
                    sheet.SetCellValue(114, "A", uwRelItem, "TitleUCCNotes");
                    sheet.SetCellValue(126, "A", uwRelItem, "EnvironmentalNotes");

                    for (int i = 0; i < 48; i++)
                    {
                        sheet.SetCellValue(18+i, "A", null);
                        sheet.SetCellValue(18 + i, "C", 0.0);
                        sheet.SetCellValue(18 + i, "D", 0.0);
                        sheet.SetCellValue(18 + i, "E", 0.0);
                        sheet.SetCellValue(18 + i, "F", 0.0);
                        sheet.SetCellValue(18 + i, "G", 0.0);
                        sheet.SetCellValue(18 + i, "H", 0.0);
                        sheet.SetCellValue(18 + i, "I", 0.0);
                        sheet.SetCellValue(18 + i, "J", 0.0);
                        sheet.SetCellValue(18 + i, "K", 0.0);
                        sheet.SetCellValue(18 + i, "L", 0.0);
                        sheet.SetCellValue(18 + i, "M", 0.0);
                        sheet.SetCellValue(18 + i, "N", 0.0);
                        sheet.SetCellValue(18 + i, "O", 0.0);
                        sheet.SetCellValue(18 + i, "P", 0.0);
                        sheet.SetCellValue(18 + i, "Q", 0.0);
                        sheet.SetCellValue(18 + i, "R", 0.0);
                    }
                    var iRow = 0;
                    foreach (var cashFlowItem in relCFdata)
                    {
                        sheet.SetCellValue(18 + iRow, "A", cashFlowItem, "CashFlowDate");
                        sheet.SetCellValue(18 + iRow, "C", cashFlowItem, "Principal");
                        sheet.SetCellValue(18 + iRow, "D", cashFlowItem, "Interest");
                        sheet.SetCellValue(18 + iRow, "E", cashFlowItem, "MiscIncome3");
                        sheet.SetCellValue(18 + iRow, "F", cashFlowItem, "MiscIncome4");
                        sheet.SetCellValue(18 + iRow, "G", cashFlowItem, "MiscIncome5");
                        sheet.SetCellValue(18 + iRow, "H", cashFlowItem, "MiscIncome6");
                        sheet.SetCellValue(18 + iRow, "I", cashFlowItem, "BackTaxes");
                        sheet.SetCellValue(18 + iRow, "J", cashFlowItem, "Legal");
                        sheet.SetCellValue(18 + iRow, "K", cashFlowItem, "Travel");
                        sheet.SetCellValue(18 + iRow, "L", cashFlowItem, "BrokerFee");
                        sheet.SetCellValue(18 + iRow, "M", cashFlowItem, "REOTax");
                        sheet.SetCellValue(18 + iRow, "N", cashFlowItem, "REOins");
                        sheet.SetCellValue(18 + iRow, "O", cashFlowItem, "CapEx");
                        sheet.SetCellValue(18 + iRow, "P", cashFlowItem, "TiLc");
                        sheet.SetCellValue(18 + iRow, "Q", cashFlowItem, "Environ");
                        sheet.SetCellValue(18 + iRow, "R", cashFlowItem, "Misc");
                        iRow++;
                    }
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
