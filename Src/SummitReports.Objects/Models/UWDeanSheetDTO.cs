using System;
using System.Collections.Generic;
using System.Text;

namespace SummitReports.Objects.Models
{
    public class UWDeanSheetDTO
    {
        //<summary></summary>
        //[Key, Column("BidPoolId", Order=1)]
        public int? BidPoolId { get; set; }
        //<summary></summary>
        //[Key, Column("uwRelationshipId", Order=2)]
        public int uwRelationshipId { get; set; }
        //<summary></summary>
        //[Key, Column("RelationshipName", Order=3)]
        public string RelationshipName { get; set; }
        //<summary></summary>
        //[Key, Column("BidPool", Order=4)]
        public string BidPool { get; set; }
        //<summary></summary>
        //[Key, Column("BidSubPoolName", Order=5)]
        public string BidSubPoolName { get; set; }
        //<summary></summary>
        //[Key, Column("UW", Order=6)]
        public string UW { get; set; }
        //<summary></summary>
        //[Key, Column("LoanCount", Order=7)]
        public int? LoanCount { get; set; }
        //<summary></summary>
        //[Key, Column("UPBSum", Order=8)]
        //[DecimalPrecision(38, 2)]
        public decimal? UPBSum { get; set; }
        //<summary></summary>
        //[Key, Column("BidAmount", Order=9)]
        //[DecimalPrecision(14, 2)]
        public decimal BidAmount { get; set; }
        //<summary></summary>
        //[Key, Column("BidUPB", Order=10)]
        public double? BidUPB { get; set; }
        //<summary></summary>
        //[Key, Column("DiscountRate", Order=11)]
        //[DecimalPrecision(10, 6)]
        public decimal DiscountRate { get; set; }
        //<summary></summary>
        //[Key, Column("TrailConC", Order=12)]
        public double? TrailConC { get; set; }
        //<summary></summary>
        //[Key, Column("ProjConC", Order=13)]
        public double? ProjConC { get; set; }
        //<summary></summary>
        //[Key, Column("MOIC", Order=14)]
        public double? MOIC { get; set; }
        //<summary></summary>
        //[Key, Column("Recovery", Order=15)]
        public double? Recovery { get; set; }
        //<summary></summary>
        //[Key, Column("AppraisalDate", Order=16)]
        public DateTime? AppraisalDate { get; set; }
        //<summary></summary>
        //[Key, Column("AppraisalValue", Order=17)]
        //[DecimalPrecision(38, 2)]
        public decimal? AppraisalValue { get; set; }
        //<summary></summary>
        //[Key, Column("BusinessAssets", Order=18)]
        //[DecimalPrecision(38, 2)]
        public decimal? BusinessAssets { get; set; }
        //<summary></summary>
        //[Key, Column("BankTotal", Order=19)]
        //[DecimalPrecision(38, 2)]
        public decimal? BankTotal { get; set; }
        //<summary></summary>
        //[Key, Column("BPOValue", Order=20)]
        //[DecimalPrecision(38, 2)]
        public decimal? BPOValue { get; set; }
        //<summary></summary>
        //[Key, Column("SIMValue", Order=21)]
        //[DecimalPrecision(38, 2)]
        public decimal? SIMValue { get; set; }
        //<summary></summary>
        //[Key, Column("BidAppr", Order=22)]
        public double? BidAppr { get; set; }
        //<summary></summary>
        //[Key, Column("BidBPO", Order=23)]
        public double? BidBPO { get; set; }
        //<summary></summary>
        //[Key, Column("BidSIMValue", Order=24)]
        public double? BidSIMValue { get; set; }
        //<summary></summary>
        //[Key, Column("PHLast3mth", Order=25)]
        //[DecimalPrecision(14, 2)]
        public decimal PHLast3mth { get; set; }
        //<summary></summary>
        //[Key, Column("PHLast6mth", Order=26)]
        //[DecimalPrecision(14, 2)]
        public decimal PHLast6mth { get; set; }
        //<summary></summary>
        //[Key, Column("PHLast9mth", Order=27)]
        //[DecimalPrecision(14, 2)]
        public decimal PHLast9mth { get; set; }
        //<summary></summary>
        //[Key, Column("PHLast12mth", Order=28)]
        //[DecimalPrecision(14, 2)]
        public decimal PHLast12mth { get; set; }
        //<summary></summary>
        //[Key, Column("Recourse", Order=29)]
        public string Recourse { get; set; }
        //<summary></summary>
        //[Key, Column("PrimaryCollateralType", Order=30)]
        public string PrimaryCollateralType { get; set; }
        //<summary></summary>
        //[Key, Column("YearBuilt", Order=31)]
        public string YearBuilt { get; set; }
        //<summary></summary>
        //[Key, Column("REUnit", Order=32)]
        public string REUnit { get; set; }
        //<summary></summary>
        //[Key, Column("REBasis", Order=33)]
        public int REBasis { get; set; }
        //<summary></summary>
        //[Key, Column("PrimaryLocation", Order=34)]
        public string PrimaryLocation { get; set; }
        //<summary></summary>
        //[Key, Column("Eyes", Order=35)]
        public string Eyes { get; set; }
        //public IDictionary<string, object> DynamicProperties { get; set; } <== uncomment to support open types

    }

}
