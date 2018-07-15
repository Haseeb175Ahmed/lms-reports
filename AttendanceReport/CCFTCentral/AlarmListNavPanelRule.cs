namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmListNavPanelRule")]
    public partial class AlarmListNavPanelRule
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid GlobalID { get; set; }

        public Guid PanelGlobalID { get; set; }

        public int? AlarmZoneID { get; set; }

        public int? SourceID { get; set; }

        public int? EventGroup { get; set; }

        public int? EventType { get; set; }

        public virtual AlarmZone AlarmZone { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        public virtual Viewer Viewer { get; set; }
    }
}
