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
    public class ModelReport : SummitReportBaseObject
    {
        public ModelReport() : base(@"ModelReport\ModelReportTemplate.xlsx")
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="uwRelationshipId"></param>
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
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("MODEL"));
                }
                this.workbook.ClearStyleCache();
                /* VERY IMPORTANT NOTE - if you mess with the template,  sometimes you may have to Build->Clear Solution before the DLL that contains your template will get refreshed */


                /* Start Your Sheet Creation Code Here */
                var standardStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, FontColor = IndexedColors.Red.AsXSSFColor(), BackgroundColor = IndexedColors.Green.AsXSSFColor() };
                var boldStyle = new XSSFNPoiStyle() { FillPattern = FillPattern.SolidForeground, FillForegroundColor = IndexedColors.PaleBlue.AsXSSFColor(), IsBold = true, VerticalAlignment = VerticalAlignment.Top, HorizontalAlignment = HorizontalAlignment.Left, WrapText = true };

                string sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_Relationship] WHERE BidPoolId=@p0;SELECT GETDATE() as ThisDate, 'SQL LITERAL' as ThisString;";
                var retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, BidPoolId);
                DataTable firstResultSet = retDataSet.Tables[0];
                foreach( DataRow row in firstResultSet.Rows)
                {
                    
                    sheet = workbook.CloneSheet(this.workbook.GetSheetIndex("MODEL"));
                    workbook.SetSheetName(workbook.NumberOfSheets-1, row["RelationshipName"].ToString());

<<<<<<< HEAD
                System.Data.DataTable secondResultSet = retDataSet.Tables[1];
                DataRow firstRow2nd = secondResultSet.Rows[0];
                sheet.SetCellValue(3, "D", "@DR2->");
                sheet.SetCellValue(3, "E", firstRow2nd, "ThisDate");
                
                var currentRow = 4;
                //var npoiBorderStyle = new XSSFNPoiStyle() { Border = CellBorder.All, FontColor = new XSSFColor(System.Drawing.Color.Red) };
                var npoiBorderStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, FontColor = IndexedColors.Red.AsXSSFColor(), FillForegroundColor = IndexedColors.Green.AsXSSFColor() };
                sheet.CreateRow(currentRow);
                sheet.SetCellValue(currentRow, "A", "@DR2->").SetCellStyle(npoiBorderStyle);                
                sheet.SetCellValue(currentRow, "E", firstRow2nd, "ThisDate").SetCellStyle(npoiBorderStyle.SetFormatStyle(FormatStyle.Date));
                sheet.SetCellValue(currentRow, "B", 999.99).SetCellStyle(npoiBorderStyle.SetFormatStyle(FormatStyle.Default));
                sheet.SetCellValue(currentRow, "C", 999.99).SetCellStyle(npoiBorderStyle.SetFormatStyle(FormatStyle.Currency));
                sheet.SetCellValue(currentRow, "D", 99999.99).SetCellStyle(npoiBorderStyle.SetFormatStyle("#,##0.0000"));
                npoiBorderStyle.FillPattern= FillPattern.SolidForeground;
                npoiBorderStyle.FillForegroundColor = IndexedColors.Pink.AsXSSFColor();
                npoiBorderStyle.IsBold = true;
                sheet.SetCellValue(currentRow, "F", firstRow2nd, "ThisDate").SetCellStyle(npoiBorderStyle.SetFormatStyle("mm/dd"));
                currentRow++;
=======
                    sheet.SetCellValue(0, "D", "@DR1->");
                    sheet.SetCellValue(0, "E", row, "uwRelationshipId").SetCellStyle(standardStyle);
                    sheet.SetCellValue(1, "D", "@DR2->");
                    sheet.SetCellValue(1, "E", row, "Underwriter").SetCellStyle(standardStyle);
                    sheet.SetCellValue(2, "D", "@DR2->").SetCellStyle(standardStyle); ;
                    sheet.SetCellValue(2, "E", row, "UPBSum").SetCellStyle(standardStyle.SetFormatStyle(FormatStyle.Currency));
                    sheet.SetCellValue(3, "D", "@DR3->").SetCellStyle(standardStyle.SetFormatStyle(FormatStyle.Default));
                    sheet.SetCellValue(3, "E", row, "CurrentStatus");
                    sheet.SetCellValue(4, "D", "@DR4->").SetCellStyle(boldStyle);
                    sheet.SetCellValue(4, "E", row, "ProFormaStatus").SetCellStyle(boldStyle);
                    sheet.SetCellValue(5, "D", "@DR5->").SetCellStyle(boldStyle);
                    sheet.SetCellValue(5, "E", row, "ExitStrategyText").SetCellStyle(boldStyle);
                    sheet.GetRow(5).Height = 1540;
                    sheet.SetColumnWidth("E", 9800);
                }
                workbook.RemoveSheetAt(this.workbook.GetSheetIndex("MODEL"));
>>>>>>> master
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
