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
                var resultSetIndex = 1;
                foreach (System.Data.DataTable table in retDataSet.Tables)
                {
                    foreach (System.Data.DataRow row in table.Rows)
                    {
                        if (resultSetIndex == 1) //first result set
                        {
                            sheet.SetCellValue(1, "D", "@DR1->");
                            sheet.SetCellValue(1, "E", row, "uwRelationshipId");
                            sheet.SetCellValue(2, "D", "@DR1->");
                            sheet.SetCellValue(2, "E", row, "RelationshipName");
                        }
                        else if (resultSetIndex == 2)
                        { //first result set 
                            sheet.SetCellValue(3, "D", "@DR2->");
                            sheet.SetCellValue(3, "E", row, "ThisDate");
                        }

                    }
                    resultSetIndex++;
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
