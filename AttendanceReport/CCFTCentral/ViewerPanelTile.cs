namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ViewerPanelTile")]
    public partial class ViewerPanelTile
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ViewerPanelTile()
        {
            AlarmDetailsTileFields = new HashSet<AlarmDetailsTileField>();
            CardholderDetailsTileFields = new HashSet<CardholderDetailsTileField>();
            EventTrailTileItems = new HashSet<EventTrailTileItem>();
            StatusTileItems = new HashSet<StatusTileItem>();
        }

        public int FTItemID { get; set; }

        public Guid PanelGlobalID { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public Guid ClassID { get; set; }

        [Required]
        [StringLength(200)]
        public string Contract { get; set; }

        public double Left { get; set; }

        public double Top { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        [Required]
        [StringLength(200)]
        public string Title { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmDetailsTileField> AlarmDetailsTileFields { get; set; }

        public virtual CardholderDetailsTile CardholderDetailsTile { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderDetailsTileField> CardholderDetailsTileFields { get; set; }

        public virtual CardholderImagesTile CardholderImagesTile { get; set; }

        public virtual EventTrailTile EventTrailTile { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventTrailTileItem> EventTrailTileItems { get; set; }

        public virtual StatusTile StatusTile { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<StatusTileItem> StatusTileItems { get; set; }

        public virtual ViewerPanel ViewerPanel { get; set; }
    }
}
