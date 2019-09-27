using System;
using System.Collections.Generic;
using System.Text;

namespace SummitReport.Infrastructure
{
    public interface ISummitReportSettings
    {
        string ConnectionString { get; set; }
        string SchemaName { get; set; }
        bool VerboseMessages { get; set; }
        string Version { get; set; }
    }
}
