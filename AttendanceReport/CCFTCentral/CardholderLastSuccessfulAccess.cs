namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardholderLastSuccessfulAccess")]
    public partial class CardholderLastSuccessfulAccess
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CardholderID { get; set; }

        public DateTime AccessTime { get; set; }

        public int ZoneID { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
