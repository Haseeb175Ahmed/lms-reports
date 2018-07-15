namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FTItemType")]
    public partial class FTItemType
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public FTItemType()
        {
            ConfigDispatchFieldControllerItemTypePairs = new HashSet<ConfigDispatchFieldControllerItemTypePair>();
            CustomFields = new HashSet<CustomField>();
            FTItems = new HashSet<FTItem>();
            FTItemTypeStates = new HashSet<FTItemTypeState>();
            IconSets = new HashSet<IconSet>();
            RemoteOverrideInfoes = new HashSet<RemoteOverrideInfo>();
            RemoteOverrideVerbInfoes = new HashSet<RemoteOverrideVerbInfo>();
            SessionPermissions = new HashSet<SessionPermission>();
            SingleControllerItems = new HashSet<SingleControllerItem>();
            UnnamedItemTypes = new HashSet<UnnamedItemType>();
            EventTypes = new HashSet<EventType>();
        }

        public int ID { get; set; }

        public Guid? ClassID { get; set; }

        public int? DefaultIconSetID { get; set; }

        public short IconStateCount { get; set; }

        public int NameKeyCode { get; set; }

        public int? LicenceLimit { get; set; }

        [Required]
        [StringLength(50)]
        public string Name { get; set; }

        public int BrandedNameKeyCode { get; set; }

        public bool AllowUserConfigIconSets { get; set; }

        [Column(TypeName = "image")]
        public byte[] IconImage { get; set; }

        public int? A2DType { get; set; }

        public Guid RemoteCLSID { get; set; }

        public bool NameMustBeUnique { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ConfigDispatchFieldControllerItemTypePair> ConfigDispatchFieldControllerItemTypePairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CustomField> CustomFields { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems { get; set; }

        public virtual IconSet IconSet { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItemTypeState> FTItemTypeStates { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<IconSet> IconSets { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RemoteOverrideInfo> RemoteOverrideInfoes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RemoteOverrideVerbInfo> RemoteOverrideVerbInfoes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SessionPermission> SessionPermissions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SingleControllerItem> SingleControllerItems { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<UnnamedItemType> UnnamedItemTypes { get; set; }

        public virtual RemoteOverrideCancelFlag RemoteOverrideCancelFlag { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventType> EventTypes { get; set; }
    }
}
