namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ActionPlan")]
    public partial class ActionPlan
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ActionPlan()
        {
            ActionPlanOutputs = new HashSet<ActionPlanOutput>();
            ActionPlanStates = new HashSet<ActionPlanState>();
            ActionPlanTriggers = new HashSet<ActionPlanTrigger>();
            EventGroups = new HashSet<EventGroup>();
            EventGroupItemActionPlans = new HashSet<EventGroupItemActionPlan>();
            EventGroupZoneActionPlans = new HashSet<EventGroupZoneActionPlan>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ActionPlanOutput> ActionPlanOutputs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ActionPlanState> ActionPlanStates { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ActionPlanTrigger> ActionPlanTriggers { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroup> EventGroups { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemActionPlan> EventGroupItemActionPlans { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupZoneActionPlan> EventGroupZoneActionPlans { get; set; }
    }
}
