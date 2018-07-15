namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EventGroupItemAlarmInstruction")]
    public partial class EventGroupItemAlarmInstruction
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SiteItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int EventGroupID { get; set; }

        public int AlarmInstructionID { get; set; }

        public virtual AlarmInstruction AlarmInstruction { get; set; }

        public virtual EventGroup EventGroup { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
