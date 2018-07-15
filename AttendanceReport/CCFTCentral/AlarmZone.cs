namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmZone")]
    public partial class AlarmZone
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public AlarmZone()
        {
            AlarmListNavPanelRules = new HashSet<AlarmListNavPanelRule>();
            AlarmZoneRelations = new HashSet<AlarmZoneRelation>();
            AlarmZoneRelations1 = new HashSet<AlarmZoneRelation>();
            EventGroupItemAlarmZones = new HashSet<EventGroupItemAlarmZone>();
            EventGroupZoneActionPlans = new HashSet<EventGroupZoneActionPlan>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Required]
        [StringLength(3)]
        public string DiallingGG { get; set; }

        [Required]
        [StringLength(5)]
        public string DiallingACCT { get; set; }

        public int AlarmDiallerID { get; set; }

        public int? PrearmTimeout { get; set; }

        public bool ResetConfAlarms { get; set; }

        public bool DisableCAOnAccess { get; set; }

        public bool FailArmUnack { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmListNavPanelRule> AlarmListNavPanelRules { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmZoneRelation> AlarmZoneRelations { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmZoneRelation> AlarmZoneRelations1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemAlarmZone> EventGroupItemAlarmZones { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupZoneActionPlan> EventGroupZoneActionPlans { get; set; }
    }
}
