namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DVRCamera")]
    public partial class DVRCamera
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? DVRTypeID { get; set; }

        [StringLength(110)]
        public string CameraUID { get; set; }

        public int? FTControllerID { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
