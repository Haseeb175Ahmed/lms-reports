namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EventType")]
    public partial class EventType
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public EventType()
        {
            EventTypeDefaultInstructions = new HashSet<EventTypeDefaultInstruction>();
            EventTypeItemAlarmInstructions = new HashSet<EventTypeItemAlarmInstruction>();
            FTItemTypes = new HashSet<FTItemType>();
            NotificationFilters = new HashSet<NotificationFilter>();
        }

        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ID { get; set; }

        public int NameCode { get; set; }

        public int MessageCode { get; set; }

        public int EventGroupID { get; set; }

        [StringLength(50)]
        public string Name { get; set; }

        public virtual EventGroup EventGroup { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventTypeDefaultInstruction> EventTypeDefaultInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventTypeItemAlarmInstruction> EventTypeItemAlarmInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItemType> FTItemTypes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<NotificationFilter> NotificationFilters { get; set; }
    }
}
