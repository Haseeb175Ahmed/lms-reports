namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardholderImagesTile")]
    public partial class CardholderImagesTile
    {
        [Key]
        public Guid TileGlobalID { get; set; }

        public bool AllPersonalData { get; set; }

        public virtual ViewerPanelTile ViewerPanelTile { get; set; }
    }
}
