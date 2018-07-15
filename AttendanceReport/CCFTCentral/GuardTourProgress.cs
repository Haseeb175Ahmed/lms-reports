namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("GuardTourProgress")]
    public partial class GuardTourProgress
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

        public DateTime MinExpectedTime { get; set; }

        public DateTime MaxExpectedTime { get; set; }

        public DateTime LateTime { get; set; }

        public DateTime? ActualTime { get; set; }

        public bool MustEnter { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual GuardTour GuardTour { get; set; }
    }
}
