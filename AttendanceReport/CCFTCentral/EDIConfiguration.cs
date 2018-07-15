namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EDIConfiguration")]
    public partial class EDIConfiguration
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int Type { get; set; }

        public int? Mode { get; set; }

        public DateTime? LastAttempt { get; set; }

        public DateTime? LastRun { get; set; }

        [StringLength(255)]
        public string Info1 { get; set; }

        [StringLength(255)]
        public string Info2 { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
