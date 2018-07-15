namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("LiftCarLevel")]
    public partial class LiftCarLevel
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int LiftCarID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int LevelNumber { get; set; }

        [Required]
        [StringLength(110)]
        public string Name { get; set; }

        public int? InputID { get; set; }

        public int? OutputID { get; set; }

        public int? AccessZoneID { get; set; }

        public virtual AccessZone AccessZone { get; set; }

        public virtual LiftCar LiftCar { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        public virtual Output Output { get; set; }
    }
}
