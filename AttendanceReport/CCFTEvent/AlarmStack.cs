namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmStack")]
    public partial class AlarmStack
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long EventID { get; set; }

        public int ControllerID { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] AlarmID { get; set; }

        public int RawSource { get; set; }

        public int EventType { get; set; }

        public byte Status { get; set; }

        public DateTime FirstOccurrenceTime { get; set; }

        public DateTime LastOccurrenceTime { get; set; }

        public int OccurrenceCount { get; set; }

        public DateTime ForgetFloodStartTime { get; set; }

        public virtual Event Event { get; set; }
    }
}
