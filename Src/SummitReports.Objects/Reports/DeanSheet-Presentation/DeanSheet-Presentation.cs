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
    public class DeanSheetPresentation : SummitReportBaseObject
    {
        public DeanSheetPresentation() : base(@"DeanSheet-Presentation\DeanSheet-Presentation.xlsx")
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

                string sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_DeanSheet] WHERE [BidPoolIdId]=@p0 ORDER BY uwRelationshipId ASC;";
                var dataArr = await MarsDb.Query<UWRelationshipDTO>(sSQL, BidPoolId);
                foreach (var data in dataArr)
                {
                    sheet.SetCellValue(4, "B", data, "RelationshipName");
                    sheet.SetCellValue(5, "F", data, "UW");
                    sheet.SetCellValue(6, "B", "LITERAL DATA");
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
