namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardLayoutASCIIMifare")]
    public partial class CardLayoutASCIIMifare
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

        [Key]
        [Column(Order = 3)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SectorNumber { get; set; }

        [Key]
        [Column(Order = 4)]
        [MaxLength(6)]
        public byte[] ReadKey { get; set; }

        [Key]
        [Column(Order = 5)]
        [MaxLength(6)]
        public byte[] WriteKey { get; set; }

        [Key]
        [Column(Order = 6)]
        [MaxLength(6)]
        public byte[] MadKey { get; set; }

        [Key]
        [Column(Order = 7)]
        public bool WriteMad { get; set; }

        [Key]
        [Column(Order = 8)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int MadCode { get; set; }

        [Key]
        [Column(Order = 9)]
        public bool EncodeAsASCII { get; set; }
    }
}
