namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("GuardTourCheckpoint")]
    public partial class GuardTourCheckpoint
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int GuardTourID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CheckpointIndex { get; set; }

        public int DoorID { get; set; }

        public int ExpectedTime { get; set; }

        public int DeviationTime { get; set; }

        public int LateTime { get; set; }

        public bool MustEnter { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual GuardTour GuardTour { get; set; }
    }
}
