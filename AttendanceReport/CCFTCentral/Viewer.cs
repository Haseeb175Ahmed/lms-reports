namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Viewer")]
    public partial class Viewer
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Viewer()
        {
            AlarmListNavPanelDisplayColumns = new HashSet<AlarmListNavPanelDisplayColumn>();
            AlarmListNavPanelRules = new HashSet<AlarmListNavPanelRule>();
            ViewerPanels = new HashSet<ViewerPanel>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public bool WorkstationRestrictionsApply { get; set; }

        public int DisplayOrder { get; set; }

        public int NavPanelDock { get; set; }

        public double DefaultDepth { get; set; }

        [Required]
        [StringLength(200)]
        public string ViewerContract { get; set; }

        [Required]
        [StringLength(200)]
        public string NavTileContract { get; set; }

        public Guid NavPanelClassID { get; set; }

        public double ViewerWidth { get; set; }

        public double ViewerHeight { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmListNavPanelDisplayColumn> AlarmListNavPanelDisplayColumns { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmListNavPanelRule> AlarmListNavPanelRules { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ViewerPanel> ViewerPanels { get; set; }
    }
}
