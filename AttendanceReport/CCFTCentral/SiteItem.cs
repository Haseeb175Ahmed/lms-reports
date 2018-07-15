namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SiteItem")]
    public partial class SiteItem
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public SiteItem()
        {
            AccessTriplets = new HashSet<AccessTriplet>();
            AlarmListNavPanelRules = new HashSet<AlarmListNavPanelRule>();
            ClubAlarmZones = new HashSet<ClubAlarmZone>();
            DVRCameras = new HashSet<DVRCamera>();
            EventGroupDefaultInstructions = new HashSet<EventGroupDefaultInstruction>();
            EventGroupItemActionPlans = new HashSet<EventGroupItemActionPlan>();
            EventGroupItemAlarmInstructions = new HashSet<EventGroupItemAlarmInstruction>();
            EventGroupItemAlarmZones = new HashSet<EventGroupItemAlarmZone>();
            EventTypeDefaultInstructions = new HashSet<EventTypeDefaultInstruction>();
            EventTypeItemAlarmInstructions = new HashSet<EventTypeItemAlarmInstruction>();
            FieldDevices = new HashSet<FieldDevice>();
            LiftCarLevels = new HashSet<LiftCarLevel>();
            SaltoAccessTriplets = new HashSet<SaltoAccessTriplet>();
            GBusURIs = new HashSet<GBusURI>();
            Overrides = new HashSet<Override>();
            ReflectIOPairs = new HashSet<ReflectIOPair>();
            SitePlanGraphics = new HashSet<SitePlanGraphic>();
            DefinedEOLResistances = new HashSet<DefinedEOLResistance>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? IconSetID { get; set; }

        public int? ScheduleID { get; set; }

        [MaxLength(1000)]
        public byte[] ConfigurationData { get; set; }

        [Required]
        [StringLength(4)]
        public string DiallingCCC { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessTriplet> AccessTriplets { get; set; }

        public virtual AccessZone AccessZone { get; set; }

        public virtual AlarmDialler AlarmDialler { get; set; }

        public virtual AlarmKeypad AlarmKeypad { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmListNavPanelRule> AlarmListNavPanelRules { get; set; }

        public virtual AlarmZone AlarmZone { get; set; }

        public virtual Bert Bert { get; set; }

        public virtual BiometricReader BiometricReader { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ClubAlarmZone> ClubAlarmZones { get; set; }

        public virtual Door Door { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DVRCamera> DVRCameras { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupDefaultInstruction> EventGroupDefaultInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemActionPlan> EventGroupItemActionPlans { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemAlarmInstruction> EventGroupItemAlarmInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemAlarmZone> EventGroupItemAlarmZones { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventTypeDefaultInstruction> EventTypeDefaultInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventTypeItemAlarmInstruction> EventTypeItemAlarmInstructions { get; set; }

        public virtual FenceController FenceController { get; set; }

        public virtual FenceZone FenceZone { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FieldDevice> FieldDevices { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        public virtual GuardTour GuardTour { get; set; }

        public virtual IconSet IconSet { get; set; }

        public virtual InterlockSetting InterlockSetting { get; set; }

        public virtual LiftCar LiftCar { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LiftCarLevel> LiftCarLevels { get; set; }

        public virtual RemoteElmo RemoteElmo { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SaltoAccessTriplet> SaltoAccessTriplets { get; set; }

        public virtual SaltoServer SaltoServer { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<GBusURI> GBusURIs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Override> Overrides { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ReflectIOPair> ReflectIOPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SitePlanGraphic> SitePlanGraphics { get; set; }

        public virtual SoftwareProcess SoftwareProcess { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DefinedEOLResistance> DefinedEOLResistances { get; set; }
    }
}
