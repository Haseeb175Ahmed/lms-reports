namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class CardLayoutMifarePlu
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

        [Required]
        [MaxLength(16)]
        public byte[] SectorKey { get; set; }

        [Required]
        [MaxLength(16)]
        public byte[] MadKey { get; set; }

        public bool WriteMad { get; set; }

        public bool CardaxIssued { get; set; }

        public bool EraseCardaxData { get; set; }

        [Required]
        [MaxLength(6)]
        public byte[] SectorKeyClassic { get; set; }

        [Required]
        [MaxLength(6)]
        public byte[] MadKeyClassic { get; set; }

        public bool UseRandomIDs { get; set; }

        public bool UseXCardsOnly { get; set; }

        public int CardUIDLength { get; set; }

        public virtual CardLayoutObject CardLayoutObject { get; set; }
    }
}
