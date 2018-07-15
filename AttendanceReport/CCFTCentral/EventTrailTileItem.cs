namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EventTrailTileItem")]
    public partial class EventTrailTileItem
    {
        [Key]
        [Column(Order = 0)]
        public Guid TileGlobalID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ItemID { get; set; }

        public virtual ViewerPanelTile ViewerPanelTile { get; set; }
    }
}
