namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardLayoutSalto")]
    public partial class CardLayoutSalto
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SideIndex { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ObjectIndex { get; set; }

        public int SectorNumber { get; set; }

        public int SectorCount { get; set; }

        public bool EventsOnCard { get; set; }

        public virtual CardLayoutObject CardLayoutObject { get; set; }
    }
}
