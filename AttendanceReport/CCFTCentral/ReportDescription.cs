namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ReportDescription")]
    public partial class ReportDescription
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int ReportType { get; set; }

        [Column(TypeName = "image")]
        public byte[] ConfigurationData { get; set; }

        public int? AccessibleByChildDivisions { get; set; }

        [Required]
        [StringLength(100)]
        public string Timezone { get; set; }

        [Required]
        [StringLength(100)]
        public string TimezonePOSIX { get; set; }

        public int LanguageID { get; set; }

        public int CustomTimezone { get; set; }

        public int Copies { get; set; }

        public int CustomLanguage { get; set; }

        public int scheduleEnabled { get; set; }

        public int scheduleOption { get; set; }

        public DateTime? scheduleFrom { get; set; }

        public DateTime? scheduleUntil { get; set; }

        public int scheduleRepeatCount { get; set; }

        public int scheduleRepeatType { get; set; }

        public int scheduleWeekdaysMask { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
