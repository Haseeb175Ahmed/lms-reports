namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ActionPlanState")]
    public partial class ActionPlanState
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ActionPlanID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int StateCode { get; set; }

        public int Priority { get; set; }

        public int EscalationTime { get; set; }

        public int EscalationActiveTime { get; set; }

        public int EscalationAlarmCount { get; set; }

        public int EscalationDoublePointTime { get; set; }

        public int EscalationDoubleZoneTime { get; set; }

        public int DeactivateOutputsWhenInactive { get; set; }

        public int AlarmDiallerID { get; set; }

        public bool KeepAlarmRelaysActive { get; set; }

        public virtual ActionPlan ActionPlan { get; set; }
    }
}
