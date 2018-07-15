namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SitePlanTileConfig")]
    public partial class SitePlanTileConfig
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid GlobalID { get; set; }

        [Key]
        [Column(Order = 2)]
        public Guid PanelGlobalID { get; set; }

        public int SiteplanConfigType { get; set; }

        public int? SpecificSitePlanItemID { get; set; }
    }
}
