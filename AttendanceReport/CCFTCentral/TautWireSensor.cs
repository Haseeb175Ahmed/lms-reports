namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("TautWireSensor")]
    public partial class TautWireSensor
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public float DetectionWindow { get; set; }

        public int DetectionThreshold { get; set; }

        public float AlarmLatchDuration { get; set; }

        public bool LED { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }
    }
}
