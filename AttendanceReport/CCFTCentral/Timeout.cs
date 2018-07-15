namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Timeout")]
    public partial class Timeout
    {
        [Key]
        [Column(Order = 0)]
        public Guid ClassID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Cookie1 { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Cookie2 { get; set; }

        public DateTime Time { get; set; }

        [Required]
        [MaxLength(32)]
        public byte[] UserData { get; set; }

        [Key]
        [Column(Order = 3)]
        public Guid GlobalID { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public int? FTItemID { get; set; }
    }
}
