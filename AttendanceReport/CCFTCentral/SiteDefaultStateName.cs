namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SiteDefaultStateName")]
    public partial class SiteDefaultStateName
    {
        [Key]
        [Column(Order = 0)]
        [StringLength(100)]
        public string AlarmZoneUser1StateName { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(100)]
        public string AlarmZoneUser2StateName { get; set; }

        [Key]
        [Column(Order = 2)]
        [StringLength(100)]
        public string InputOpenStateName { get; set; }

        [Key]
        [Column(Order = 3)]
        [StringLength(100)]
        public string InputCloseStateName { get; set; }

        [Key]
        [Column(Order = 4)]
        [StringLength(100)]
        public string InputBreakStateName { get; set; }

        [Key]
        [Column(Order = 5)]
        [StringLength(100)]
        public string InputShortStateName { get; set; }

        [Key]
        [Column(Order = 6)]
        [StringLength(100)]
        public string InputTamperStateName { get; set; }

        [Key]
        [Column(Order = 7)]
        [StringLength(100)]
        public string OutputOnStateName { get; set; }

        [Key]
        [Column(Order = 8)]
        [StringLength(100)]
        public string OutputOffStateName { get; set; }
    }
}
