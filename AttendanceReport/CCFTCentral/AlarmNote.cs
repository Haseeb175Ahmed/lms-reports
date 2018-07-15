namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class AlarmNote
    {
        [Key]
        [Column("AlarmNote", Order = 0)]
        [StringLength(255)]
        public string AlarmNote1 { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SortOrder { get; set; }
    }
}
