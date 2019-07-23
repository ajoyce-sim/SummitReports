using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SummitReports.Objects.Services
{
    public class UWDataService
    {
        public async Task<IEnumerable<UWRelationshipDTO>> FetchUWRelationshipData(int uwRelationshipId)
        {
            string sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_Relationship] WHERE [uwRelationshipId]=@p0;";
            return await MarsDb.Query<UWRelationshipDTO>(sSQL, uwRelationshipId);
        }

        public async Task<IEnumerable<UWRelationshipCashFlowDTO>> FetchUWRelationshipCashFlowsData(int uwRelationshipId)
        {
            string sSQL = @"SET ANSI_WARNINGS OFF; SELECT * FROM [UW].[vw_RelationshipCashFlow] WHERE [uwRelationshipId]=@p0 ORDER BY CashFlowDate ASC;";
            return await MarsDb.Query<UWRelationshipCashFlowDTO>(sSQL, uwRelationshipId);
        }
    }
}
