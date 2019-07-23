using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MARS.Entities.Models.Custom
{
    /// <summary></summary>
    public partial class UWRelationshipCashFlowDTO 
    {
        public UWRelationshipCashFlowDTO()
        {

        }
        ///<summary>Provide the actual method AfterConstructor() in a partial class and it will get called as the last point in construction.</summary>
        partial void AfterConstructor();

        //<summary></summary>
        //[Column("uwRelationshipId")]
        public int uwRelationshipId { get; set; }
        //<summary></summary>
        //[Column("RelationshipName")]
        public string RelationshipName { get; set; }
        //<summary></summary>
        //[Column("CashFlowNum")]
        public System.Int64? CashFlowNum { get; set; }
        //<summary></summary>
        //[Column("CashFlowDate")]
        public DateTime? CashFlowDate { get; set; }
        //<summary></summary>
        //[Column("Id")]
        public int Id { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("Principal")]
        public decimal Principal { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("Interest")]
        public decimal Interest { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("MiscIncome3")]
        public decimal MiscIncome3 { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("MiscIncome4")]
        public decimal MiscIncome4 { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("MiscIncome5")]
        public decimal MiscIncome5 { get; set; }
        //<summary></summary>
        ////[DecimalPrecision(14, 2)]
        //[Column("MiscIncome6")]
        public decimal MiscIncome6 { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("BackTaxes")]
        public decimal BackTaxes { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("Legal")]
        public decimal Legal { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("Travel")]
        public decimal Travel { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("BrokerFee")]
        public decimal BrokerFee { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("REOTax")]
        public decimal REOTax { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("REOins")]
        public decimal REOins { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("CapEx")]
        public decimal CapEx { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("TiLc")]
        public decimal TiLc { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("Environ")]
        public decimal Environ { get; set; }
        //<summary></summary>
        //[DecimalPrecision(14, 2)]
        //[Column("Misc")]
        public decimal Misc { get; set; }
        //<summary></summary>
        //[DecimalPrecision(19, 2)]
        //[Column("TotalIncome")]
        public decimal? TotalIncome { get; set; }
        //<summary></summary>
        //[DecimalPrecision(23, 2)]
        //[Column("TotalExpense")]
        public decimal? TotalExpense { get; set; }
        //<summary></summary>
        //[DecimalPrecision(24, 2)]
        //[Column("GrossCashFlow")]
        public decimal? GrossCashFlow { get; set; }
        //<summary></summary>
        //[DecimalPrecision(38, 2)]
        //[Column("WALNumerator")]
        public decimal? WALNumerator { get; set; }
        //<summary></summary>
        //[DecimalPrecision(15, 2)]
        //[Column("CashOnCashNumerator")]
        public decimal? CashOnCashNumerator { get; set; }
        //<summary></summary>
        //[Column("AddDate")]
        public DateTime? AddDate { get; set; }
        //<summary></summary>
        //[Column("AddUser")]
        public string AddUser { get; set; }
        //public IDictionary<string, object> DynamicProperties { get; set; } <== uncomment to support open types
    }
}