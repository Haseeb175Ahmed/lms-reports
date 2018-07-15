namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EventTrailTile")]
    public partial class EventTrailTile
    {
        [Key]
        public Guid TileGlobalID { get; set; }

        public short ConfigType { get; set; }

        public short LimitType { get; set; }

        public int LimitValue { get; set; }

        public virtual ViewerPanelTile ViewerPanelTile { get; set; }
    }
}
