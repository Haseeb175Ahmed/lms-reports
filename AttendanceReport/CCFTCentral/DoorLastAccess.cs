namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DoorLastAccess")]
    public partial class DoorLastAccess
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int DoorID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AccessType { get; set; }

        public int CardholderID { get; set; }

        public DateTime AccessTime { get; set; }

        public int ZoneID { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] CardNumber { get; set; }

        public int CardFacilityCode { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
