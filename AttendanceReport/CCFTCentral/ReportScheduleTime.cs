namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ReportScheduleTime")]
    public partial class ReportScheduleTime
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ReportDescriptionID { get; set; }

        [Key]
        [Column(Order = 1)]
        public DateTime ScheduledTime { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
