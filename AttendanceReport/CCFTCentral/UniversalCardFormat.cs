namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("UniversalCardFormat")]
    public partial class UniversalCardFormat
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? WiegandFormat { get; set; }

        public int? CardNumberOffset { get; set; }

        public int? WGDNumOfBits { get; set; }

        public int? WGDFirstBitToSumEven { get; set; }

        public int? WGDBitsToSumEven { get; set; }

        public int? WGDFirstBitToSumOdd { get; set; }

        public int? WGDBitsToSumOdd { get; set; }

        public int? WGDFirstBitFacility { get; set; }

        public int? WGDNumBitsFacility { get; set; }

        public int? WGDFirstBitCardNum { get; set; }

        public int? WGDNumBitsCardNum { get; set; }

        public int? WGDFirstBitIssue { get; set; }

        public int? WGDNumBitsIssue { get; set; }

        public int? WGDStartPINOffset { get; set; }

        public int? ABAMaxDigits { get; set; }

        public int? ABAMinDigits { get; set; }

        public int? ABAFirstDigitFacility { get; set; }

        public int? ABANumDigitsFacility { get; set; }

        public int? ABAFirstDigitCardNum { get; set; }

        public int? ABANumDigitsCardNum { get; set; }

        public int? ABAFirstDigitIssue { get; set; }

        public int? ABANumDigitsIssue { get; set; }

        public int FacilityCode { get; set; }

        public int IssueOffset { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
