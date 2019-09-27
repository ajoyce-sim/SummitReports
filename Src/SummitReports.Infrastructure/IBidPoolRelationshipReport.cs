using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SummitReports.Infrastructure
{
    public interface IBidPoolRelationshipReport : ISummitReport
    {
        Task<string> BidPoolGenerateAsync(int BidPoolId);
        Task<string> RelationshipGenerateAsync(int uwRElationshipId);
    }
}
