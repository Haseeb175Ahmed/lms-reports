namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DisturbanceSensor")]
    public partial class DisturbanceSensor
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int TiltThreshold { get; set; }

        public int DynamicThreshold { get; set; }

        public int DynamicCountThreshold { get; set; }

        public float DynamicHoldTime { get; set; }

        public float DynamicMemoryLength { get; set; }

        public int MaximumEventCount { get; set; }

        public float AlarmLatchDuration { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }
    }
}
