namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Card")]
    public partial class Card
    {
        [Required]
        [StringLength(1024)]
        public string EncodedNumber { get; set; }

        public int CardTypeID { get; set; }

        public int CardholderID { get; set; }

        public int IssueLevel { get; set; }

        public int IsEnabled { get; set; }

        public DateTime? ActivationTime { get; set; }

        public DateTime? ExpiryTime { get; set; }

        public int IsTraceOn { get; set; }

        public int IsResident { get; set; }

        public int FacilityCode { get; set; }

        public DateTime? InactivityStart { get; set; }

        public DateTime? LastEncodedOrPrintedTime { get; set; }

        public int? LastEncodedOrPrintedIssueLevel { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public int ConflictCount { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] EncodedBinary { get; set; }

        public virtual Cardholder Cardholder { get; set; }

        public virtual CardType CardType { get; set; }
    }
}
