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
                    this.sheet = this.workbook.GetSheetAt(this.workbook.GetSheetIndex("Sheet1"));
                }
                this.workbook.ClearStyleCache();

                /* Using Model Specified */ 
                string sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_RelationshipCashFlow] WHERE [uwRelationshipId]=@p0 ORDER BY CashFlowDate ASC;";
                var dataArr = await MarsDb.Query<UWRelationshipDTO>(sSQL, uwRelationshipId);
                foreach (var data in dataArr)
                {
                    sheet.SetCellValue(1, "B", data, "uwRelationshipId");
                    sheet.SetCellValue(2, "B", data, "Underwriter");
                    sheet.SetCellValue(3, "B", "LITERAL DATA");
                }

                /* Using ADO.NET Specified */
                string sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_RelationshipCashFlow] WHERE [uwRelationshipId]=@p0 ORDER BY CashFlowDate ASC;SELECT GETDATE() as ThisDate, 'SQL LITERAL' as ThisString;";
                var retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, uwRelationshipId);
                DataTable firstResultSet = retDataSet.Tables[0];
                DataRow firstRow = firstResultSet.Rows[0];
                sheet.SetCellValue(1, "D", "@DR1->");
                sheet.SetCellValue(1, "E", firstRow, "uwRelationshipId");
                sheet.SetCellValue(2, "D", "@DR1->");
                sheet.SetCellValue(2, "E", firstRow, "RelationshipName");


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
