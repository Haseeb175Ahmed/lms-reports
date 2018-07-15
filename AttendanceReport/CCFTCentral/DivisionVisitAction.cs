namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DivisionVisitAction")]
    public partial class DivisionVisitAction
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        public int ID { get; set; }

        public int Action { get; set; }

        public bool RequiredAtOnsite { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
