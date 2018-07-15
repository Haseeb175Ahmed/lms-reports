namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EventGroupZoneActionPlan")]
    public partial class EventGroupZoneActionPlan
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AlarmZoneID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int EventGroupID { get; set; }

        public int ActionPlanID { get; set; }

        public virtual ActionPlan ActionPlan { get; set; }

        public virtual AlarmZone AlarmZone { get; set; }

        public virtual EventGroup EventGroup { get; set; }
    }
}
