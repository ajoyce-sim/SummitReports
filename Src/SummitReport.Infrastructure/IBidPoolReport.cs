using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SummitReport.Infrastructure
{
    public interface IBidPoolReport : ISummitReport
    {
        Task<string> GenerateAsync(int BidPoolId);
    }
}
