using System;

namespace MARS.Entities.Models.Custom
{

    public class UWRelationshipDTO
    {
        public UWRelationshipDTO()
        {
        }
        //<summary></summary>
        //[Column("uwRelationshipId")]
        public int uwRelationshipId { get; set; }
        //<summary></summary>
        //[Column("RelationshipName")]
        public string RelationshipName { get; set; }
        //<summary></summary>
        //[Column("BidPoolId")]
        public int? BidPoolId { get; set; }
        //<summary></summary>
        //[Column("BidPoolName")]
        public string BidPoolName { get; set; }
        //<summary></summary>
        //[Column("BidSubPoolId")]
        public int? BidSubPoolId { get; set; }
        //<summary></summary>
        //[Column("BidSubPoolName")]
        public string BidSubPoolName { get; set; }
        //<summary></summary>
        //[Column("LoanCount")]
        public int? LoanCount { get; set; }
        //<summary></summary>
        //[Column("RecourseFlag")]
        public bool RecourseFlag { get; set; }
        //<summary></summary>
        //[Column("SiteVisitFlag")]
        public bool SiteVisitFlag { get; set; }
        //<summary></summary>
        //[Column("UnderwriterUid")]
        public int? UnderwriterUid { get; set; }
        //<summary></summary>
        //[Column("Underwriter")]
        public string Underwriter { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("UPBSum")]
        public decimal? UPBSum { get; set; }
        //<summary></summary>
        //[Column("PrimaryLienPositionId")]
        public int PrimaryLienPositionId { get; set; }
        //<summary></summary>
        //[Column("PrimaryLienPoisition")]
        public string PrimaryLienPoisition { get; set; }
        //<summary></summary>
        //[Column("CurrentStatusId")]
        public int? CurrentStatusId { get; set; }
        //<summary></summary>
        //[Column("CurrentStatus")]
        public string CurrentStatus { get; set; }
        //<summary></summary>
        //[Column("ProFormaStatusId")]
        public int? ProFormaStatusId { get; set; }
        //<summary></summary>
        //[Column("ProFormaStatus")]
        public string ProFormaStatus { get; set; }
        //<summary></summary>
        //[DecimalPrecision(10, 6)]
        //[Column("PerformingRate")]
        public decimal? PerformingRate { get; set; }
        //<summary></summary>
        //[Column("ExitTypeId")]
        public int? ExitTypeId { get; set; }
        //<summary></summary>
        //[Column("ExitTypeDesc")]
        public string ExitTypeDesc { get; set; }
        //<summary></summary>
        //[Column("ExitStrategyText")]
        public string ExitStrategyText { get; set; }
        //<summary></summary>
        //[Column("PrimaryAddress")]
        public string PrimaryAddress { get; set; }
        //<summary></summary>
        //[Column("PrimaryCity")]
        public string PrimaryCity { get; set; }
        //<summary></summary>
        //[Column("PrimaryState")]
        public string PrimaryState { get; set; }
        //<summary></summary>
        //[Column("PrimaryZipCode")]
        public string PrimaryZipCode { get; set; }
        //<summary></summary>
        //[Column("PrimaryCounty")]
        public string PrimaryCounty { get; set; }
        //<summary></summary>
        //[Column("CollateralDescText")]
        public string CollateralDescText { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("SIMValue")]
        public decimal? SIMValue { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("BPOValueCRE")]
        public decimal? BPOValueCRE { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("AppraisalValue")]
        public decimal? AppraisalValue { get; set; }
        //<summary></summary>
        //[Column("AppraisalDate")]
        public DateTime? AppraisalDate { get; set; }
        //<summary></summary>
        //[Column("SiteVisit")]
        public DateTime? SiteVisit { get; set; }
        //<summary></summary>
        //[Column("YearBuilt")]
        public string YearBuilt { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("TotalSIM")]
        public decimal? TotalSIM { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("TotalApprBook")]
        public decimal? TotalApprBook { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("BidAmount")]
        public decimal BidAmount { get; set; }
        //<summary></summary>
        //[Column("BidUPB")]
        public double? BidUPB { get; set; }
        //<summary></summary>
        //[Column("BidSIM")]
        public double? BidSIM { get; set; }
        //<summary></summary>
        //[Column("Recovery")]
        public double? Recovery { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("TotalIncome")]
        public decimal? TotalIncome { get; set; }
        //<summary></summary>
        //[Column("MOIC")]
        public double? MOIC { get; set; }
        //<summary></summary>
        //[Column("WAL")]
        public double? WAL { get; set; }
        //<summary></summary>
        //[Column("WALNumerator")]
        public double? WALNumerator { get; set; }
        //<summary></summary>
        //[Column("GrossCashFlow")]
        public double? GrossCashFlow { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("NetCashFlow")]
        public decimal? NetCashFlow { get; set; }
        //<summary></summary>
        //[Column("LegalGross")]
        public double? LegalGross { get; set; }
        //<summary></summary>
        //[Column("BidMetric")]
        public double? BidMetric { get; set; }
        //<summary></summary>
        //[Column("SIMMetric")]
        public double? SIMMetric { get; set; }
        //<summary></summary>
        //[DecimalPrecision(10, 6)]
        //[Column("DiscountRate")]
        public decimal DiscountRate { get; set; }
        //<summary></summary>
        //[Column("PrimaryCollateralTypeId")]
        public int? PrimaryCollateralTypeId { get; set; }
        //<summary></summary>
        //[Column("PrimaryCollateralType")]
        public string PrimaryCollateralType { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("CashOnCashNumerator")]
        public decimal? CashOnCashNumerator { get; set; }
        //<summary></summary>
        //[Column("CashOnCash")]
        public double? CashOnCash { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 8)]
        //[Column("NRESIM")]
        public decimal? NRESIM { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("NREBook")]
        public decimal? NREBook { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("PHLast3mth")]
        public decimal PHLast3mth { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("PHLast6mth")]
        public decimal PHLast6mth { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("PHLast9mth")]
        public decimal PHLast9mth { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("PHLast12mth")]
        public decimal PHLast12mth { get; set; }
        //<summary></summary>
        //[Column("AssetNotes")]
        public string AssetNotes { get; set; }
        //<summary></summary>
        //[Column("CollateralValuationNotes")]
        public string CollateralValuationNotes { get; set; }
        //<summary></summary>
        //[Column("TitleUCCNotes")]
        public string TitleUCCNotes { get; set; }
        //<summary></summary>
        //[Column("EnvironmentalNotes")]
        public string EnvironmentalNotes { get; set; }
        //<summary></summary>
        //[Column("Size")]
        public int Size { get; set; }
        //<summary></summary>
        //[Column("SizeMetricId")]
        public int? SizeMetricId { get; set; }
        //<summary></summary>
        //[Column("SizeMetricDesc")]
        public string SizeMetricDesc { get; set; }
        //<summary></summary>
        //[Column("CFStartDate")]
        public DateTime? CFStartDate { get; set; }
        //<summary></summary>
        //[Column("CashOnCashStartId")]
        public int? CashOnCashStartId { get; set; }
        //<summary></summary>
        //[Column("MiscIncome3Label")]
        public string MiscIncome3Label { get; set; }
        //<summary></summary>
        //[Column("MiscIncome4Label")]
        public string MiscIncome4Label { get; set; }
        //<summary></summary>
        //[Column("MiscIncome5Label")]
        public string MiscIncome5Label { get; set; }
        //<summary></summary>
        //[Column("MiscIncome6Label")]
        public string MiscIncome6Label { get; set; }
        //<summary></summary>
        //[Column("DropFlag")]
        public bool DropFlag { get; set; }
        //<summary></summary>
        //[Column("DealResultId")]
        public int DealResultId { get; set; }
        //<summary></summary>
        //[Column("DealResultDesc")]
        public string DealResultDesc { get; set; }
        //<summary></summary>
        //[Column("AddDate")]
        public DateTime? AddDate { get; set; }
        //<summary></summary>
        //[Column("AddUser")]
        public string AddUser { get; set; }
        //public IDictionary<string, object> DynamicProperties { get; set; } <== uncomment to support open types
    }
}
