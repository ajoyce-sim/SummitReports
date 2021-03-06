﻿using System;
using NPOI.SS.UserModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using System.Data;
using SummitReports.Infrastructure;

namespace SummitReports.Objects
{
    public class ModelReport : SummitExcelReportBaseObject, IGenericReport
    {
        public ModelReport() : base(@"ModelReport\ModelReportTemplate.xlsx")
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="uwRelationshipId"></param>
        /// <returns>Name of the file generated</returns>
        public async Task<string> GenerateAsync(int Id)
        {
            try
            {
                if (!this.ReloadTemplate("MODEL")) throw new Exception("Template could not be loaded :(");
                /* Start Your Sheet Creation Code Here */
                var standardStyle = new XSSFNPoiStyle() { Border = CellBorder.All, BorderStyle = BorderStyle.Thin, FontColor = IndexedColors.Red.AsXSSFColor(), BackgroundColor = IndexedColors.Green.AsXSSFColor() };
                var boldStyle = new XSSFNPoiStyle() { FillPattern = FillPattern.SolidForeground, FillForegroundColor = IndexedColors.PaleBlue.AsXSSFColor(), IsBold = true, VerticalAlignment = VerticalAlignment.Top, HorizontalAlignment = HorizontalAlignment.Left, WrapText = true };

                string sSQL2 = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_Relationship] WHERE BidPoolId=@p0;SELECT GETDATE() as ThisDate, 'SQL LITERAL' as ThisString;";
                var retDataSet = await MarsDb.QueryAsDataSetAsync(sSQL2, Id);
                DataTable firstResultSet = retDataSet.Tables[0];
                foreach (DataRow row in firstResultSet.Rows)
                {

                    sheet = workbook.CloneSheet(this.workbook.GetSheetIndex("MODEL"));
                    workbook.SetSheetName(workbook.NumberOfSheets - 1, row["RelationshipName"].ToString().AsSheetName());

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