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
    public class SampleReport : SummitReportBaseObject
    {
        public SampleReport() : base(@"SampleReport\SampleReportTemplate.xlsx")
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

                string sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_DeanSheet] WHERE [BidPoolId]=@p0 ORDER BY BidPoolId ASC;";
                var dataArr = await MarsDb.Query<UWDeanSheetDTO>(sSQL, BidPoolId);
                foreach (var data in dataArr)
                {
                    sheet.SetCellValue(1, "A", "Portfolio Bid Worksheet);
                    sheet.SetCellValue(2, "B", 777d);
                    sheet.SetCellValue(3, "B", "LITERAL DATA Z");
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
