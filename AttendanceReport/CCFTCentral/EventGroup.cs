namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EventGroup")]
    public partial class EventGroup
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public EventGroup()
        {
            EventGroupDefaultInstructions = new HashSet<EventGroupDefaultInstruction>();
            EventGroupItemActionPlans = new HashSet<EventGroupItemActionPlan>();
            EventGroupItemAlarmInstructions = new HashSet<EventGroupItemAlarmInstruction>();
            EventGroupZoneActionPlans = new HashSet<EventGroupZoneActionPlan>();
            EventTypes = new HashSet<EventType>();
            NotificationFilters = new HashSet<NotificationFilter>();
        }

        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ID { get; set; }

        public int NameCode { get; set; }

        public int DefaultActionPlanID { get; set; }

        public int? DefaultAlarmInstructionID { get; set; }

        [StringLength(50)]
        public string Name { get; set; }

        public int DiallingXYZ { get; set; }

        public bool DiallingXYZConfigurable { get; set; }

        public virtual ActionPlan ActionPlan { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupDefaultInstruction> EventGroupDefaultInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemActionPlan> EventGroupItemActionPlans { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemAlarmInstruction> EventGroupItemAlarmInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupZoneActionPlan> EventGroupZoneActionPlans { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventType> EventTypes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<NotificationFilter> NotificationFilters { get; set; }
    }
}
