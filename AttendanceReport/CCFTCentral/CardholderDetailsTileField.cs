namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardholderDetailsTileField")]
    public partial class CardholderDetailsTileField
    {
        [Key]
        [Column(Order = 0)]
        public Guid TileGlobalID { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(100)]
        public string BrowseName { get; set; }

        public virtual ViewerPanelTile ViewerPanelTile { get; set; }
    }
}
