namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("UnlinkedAck")]
    public partial class UnlinkedAck
    {
        [Key]
        [Column(Order = 0)]
        public DateTime AckTime { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AckSourceID { get; set; }

        [Key]
        [Column(Order = 2)]
        public DateTime AlarmTime { get; set; }

        [Key]
        [Column(Order = 3)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AlarmSourceID { get; set; }

        [Key]
        [Column(Order = 4)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AlarmType { get; set; }

        [Key]
        [Column(Order = 5)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AlarmID { get; set; }

        [Key]
        [Column(Order = 6)]
        public bool AlarmActive { get; set; }

        [Key]
        [Column(Order = 7)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CardholderID { get; set; }

        [Key]
        [Column(Order = 8)]
        [StringLength(110)]
        public string CardholderName { get; set; }
    }
}
