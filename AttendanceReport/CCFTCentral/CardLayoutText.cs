namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardLayoutText")]
    public partial class CardLayoutText
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

        public int FontFlags { get; set; }

        public int FontColour { get; set; }

        public int FontSize { get; set; }

        public int FontAlignment { get; set; }

        public double FontRotation { get; set; }

        public byte FontCharSet { get; set; }

        [Required]
        [StringLength(32)]
        public string FontName { get; set; }

        public virtual CardLayoutObject CardLayoutObject { get; set; }
    }
}
