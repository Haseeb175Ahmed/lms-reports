namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmUpdate")]
    public partial class AlarmUpdate
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long AlarmEventID { get; set; }

        public long Sequence { get; set; }

        public int SequenceOrder { get; set; }

        public virtual Event Event { get; set; }
    }
}
