namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AccessZone")]
    public partial class AccessZone
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public AccessZone()
        {
            AccessZone1 = new HashSet<AccessZone>();
            AccessZoneCompetencyPairs = new HashSet<AccessZoneCompetencyPair>();
            AperioDoors = new HashSet<AperioDoor>();
            LiftCarLevels = new HashSet<LiftCarLevel>();
            Overrides = new HashSet<Override>();
            PersonalisedActions = new HashSet<PersonalisedAction>();
            Doors = new HashSet<Door>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? ParentZoneID { get; set; }

        public int DualAuthFirstClubID { get; set; }

        public int DualAuthSecondClubID { get; set; }

        public int BertID { get; set; }

        public int ZoneCountOverrides { get; set; }

        public int ZoneCountEnable { get; set; }

        public int ArmOnEmpty { get; set; }

        public int ZeroIsMidrange { get; set; }

        public int ZoneCountMin { get; set; }

        public int ZoneCountMax { get; set; }

        public int MinGrace { get; set; }

        public int MidGrace { get; set; }

        public int MaxGrace { get; set; }

        [Required]
        [StringLength(100)]
        public string MinMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string MidMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string MaxMessage { get; set; }

        public int AntiPassbackEnabled { get; set; }

        public int AntiTailgateEnabled { get; set; }

        public int AntiPassbackUseBert { get; set; }

        public int ButtonOverrideTimeout { get; set; }

        [StringLength(2048)]
        public string NoCompetenciesORMsg { get; set; }

        public int? CompetencyAccessOperation { get; set; }

        public bool PATrigger { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessZone> AccessZone1 { get; set; }

        public virtual AccessZone AccessZone2 { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessZoneCompetencyPair> AccessZoneCompetencyPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AperioDoor> AperioDoors { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LiftCarLevel> LiftCarLevels { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Override> Overrides { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalisedAction> PersonalisedActions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Door> Doors { get; set; }
    }
}
