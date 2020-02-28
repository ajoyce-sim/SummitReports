using SummitReports.Infrastructure;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace SummitReports.Objects
{
    public interface ISummitDocument
    {
        void ReplaceFieldValue(DataRow data, string ColumnName);
        void ReplaceFieldValue(DataRow data, string ColumnName, string Format);
        void ReplaceFieldValue(string ColumnName, string valueToSet);
    }

    public interface ISummitReportInternal : ISummitReport
    {
        ISummitDocument Document { get; set; }
    }
}
