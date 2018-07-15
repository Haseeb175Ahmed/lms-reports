namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ReservedCardNumber")]
    public partial class ReservedCardNumber
    {
        [Key]
        [Column(Order = 0)]
        [MaxLength(256)]
        public byte[] EncodedNumber { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FacilityCode { get; set; }

        public int CardholderID { get; set; }

        public int CardTypeID { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }
    }
}
