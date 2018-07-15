namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Licence")]
    public partial class Licence
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int UsedGraceTime { get; set; }

        [Key]
        [Column(Order = 1)]
        public DateTime LastUpdate { get; set; }

        [Key]
        [Column(Order = 2, TypeName = "timestamp")]
        [MaxLength(8)]
        [Timestamp]
        public byte[] Revision { get; set; }
    }
}
