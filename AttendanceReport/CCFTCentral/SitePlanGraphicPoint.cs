namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SitePlanGraphicPoint")]
    public partial class SitePlanGraphicPoint
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SitePlanGraphicID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SequenceNumber { get; set; }

        public int x { get; set; }

        public int y { get; set; }

        public virtual SitePlanGraphic SitePlanGraphic { get; set; }
    }
}
