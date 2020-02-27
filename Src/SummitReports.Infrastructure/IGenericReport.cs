using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SummitReports.Infrastructure
{
    public interface IGenericReport : ISummitReport
    {
        Task<string> GenerateAsync(int Id);

    }

    public interface ISummitReport
    {
        string ReportWorkPath { get; set; }
        void Clear();
        bool SaveToFile(string FileName);
        bool ReloadTemplate(string initial = "");
    }
}
