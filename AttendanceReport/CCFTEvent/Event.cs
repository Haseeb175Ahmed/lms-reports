namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Event")]
    public partial class Event
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Event()
        {
            AlarmHistories = new HashSet<AlarmHistory>();
            NotificationAlarms = new HashSet<NotificationAlarm>();
            RelatedItems = new HashSet<RelatedItem>();
        }

        public long ID { get; set; }

        public DateTime OccurrenceTime { get; set; }

        public DateTime ArrivalTime { get; set; }

        public int EventType { get; set; }

        public int DivisionID { get; set; }

        public byte Priority { get; set; }

        public byte EventClass { get; set; }

        public byte Status { get; set; }

        public byte OccurrenceCount { get; set; }

        [Required]
        [StringLength(1024)]
        public string Message { get; set; }

        [Required]
        public string Details { get; set; }

        public int? ArchiveFileID { get; set; }

        public int HomeServerID { get; set; }

        public Guid GlobalID { get; set; }

        public virtual AlarmStack AlarmStack { get; set; }

        public virtual AlarmUpdate AlarmUpdate { get; set; }

        public virtual CardEvent CardEvent { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmHistory> AlarmHistories { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<NotificationAlarm> NotificationAlarms { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RelatedItem> RelatedItems { get; set; }
    }
}
