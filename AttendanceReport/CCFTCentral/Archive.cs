namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Archive")]
    public partial class Archive
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public DateTime? LastGood { get; set; }

        public DateTime? LastTry { get; set; }

        public int LastResult { get; set; }

        [Required]
        [StringLength(260)]
        public string ArchivePaths1 { get; set; }

        [Required]
        [StringLength(260)]
        public string ArchivePaths2 { get; set; }

        [Required]
        [StringLength(260)]
        public string RestorePath { get; set; }

        [Required]
        [StringLength(50)]
        public string AccountName { get; set; }

        public bool ScheduleEnabled { get; set; }

        public int ScheduleInterval { get; set; }

        public DateTime ScheduleStart { get; set; }

        public bool EventArchiveEnabled { get; set; }

        public int EventOnlineDays { get; set; }

        public int EventOnlineMegabytes { get; set; }

        public bool ImageArchiveEnabled { get; set; }

        public int ImageOnlineDays { get; set; }

        public int ImageOnlineMegabytes { get; set; }

        [MaxLength(100)]
        public byte[] EncryptedAccountPassword { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
