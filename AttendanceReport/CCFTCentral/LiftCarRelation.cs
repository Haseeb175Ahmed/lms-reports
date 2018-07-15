namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("LiftCarRelation")]
    public partial class LiftCarRelation
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int LiftCarID { get; set; }

        public int RelatedItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int RelationCode { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual LiftCar LiftCar { get; set; }
    }
}
