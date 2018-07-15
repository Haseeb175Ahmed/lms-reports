namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Commander")]
    public partial class Commander
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? CmdrServerID { get; set; }

        public int? SerialNumber { get; set; }

        public int? IPAddress { get; set; }

        public int? IPPort { get; set; }

        public int? PollingInterval { get; set; }

        public int? PollingTimeout { get; set; }

        public DateTime? BackupLastGood { get; set; }

        public DateTime? BackupLastTry { get; set; }

        public int? BackupLastResult { get; set; }

        public DateTime? RestoreLastGood { get; set; }

        public DateTime? RestoreLastTry { get; set; }

        public int? RestoreLastResult { get; set; }

        [Column(TypeName = "image")]
        public byte[] Nonvols { get; set; }
    }
}
