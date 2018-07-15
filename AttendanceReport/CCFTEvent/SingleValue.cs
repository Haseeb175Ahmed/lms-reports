namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SingleValue
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ELID { get; set; }

        public int? Version { get; set; }

        public int? UpgradeProgress { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long NextArchiveID { get; set; }

        public DateTime? LastArchivedTime { get; set; }

        [Key]
        [Column(Order = 2)]
        public bool DatabaseAdminEnabled { get; set; }
    }
}
