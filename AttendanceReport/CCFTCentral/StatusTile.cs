namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("StatusTile")]
    public partial class StatusTile
    {
        [Key]
        public Guid TileGlobalID { get; set; }

        public short ConfigType { get; set; }

        public short IconSize { get; set; }

        public virtual ViewerPanelTile ViewerPanelTile { get; set; }
    }
}
