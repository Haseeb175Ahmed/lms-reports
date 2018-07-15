namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardLayoutGraphic")]
    public partial class CardLayoutGraphic
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

        public int TransparentColour { get; set; }

        public int Transparency { get; set; }

        public double AspectRatio { get; set; }

        public int BandWidth { get; set; }

        public byte[] Image { get; set; }

        public virtual CardLayoutObject CardLayoutObject { get; set; }
    }
}
