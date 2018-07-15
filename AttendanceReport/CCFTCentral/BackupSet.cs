namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("BackupSet")]
    public partial class BackupSet
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int BackupID { get; set; }

        [StringLength(300)]
        public string BackupPath { get; set; }

        [StringLength(200)]
        public string BackupName { get; set; }

        [StringLength(500)]
        public string BackupDescription { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int BackupType { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ScheduleType { get; set; }

        public DateTime? ScheduleTime { get; set; }

        [Key]
        [Column(Order = 3)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ValidateBackup { get; set; }

        public int? DayOfMonth { get; set; }
    }
}
