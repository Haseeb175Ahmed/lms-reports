namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmHistory")]
    public partial class AlarmHistory
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long EventID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int OperatorID { get; set; }

        [Key]
        [Column(Order = 2)]
        public DateTime Time { get; set; }

        [Key]
        [Column(Order = 3)]
        public byte Action { get; set; }

        [Key]
        [Column(Order = 4)]
        [StringLength(110)]
        public string OperatorName { get; set; }

        [Key]
        [Column(Order = 5)]
        [StringLength(1024)]
        public string Comment { get; set; }

        [Key]
        [Column(Order = 6)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SourceID { get; set; }

        public virtual Event Event { get; set; }
    }
}
