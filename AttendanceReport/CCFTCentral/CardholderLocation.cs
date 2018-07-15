namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardholderLocation")]
    public partial class CardholderLocation
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CardholderID { get; set; }

        public DateTime AccessTime { get; set; }

        public int AccessType { get; set; }

        public int DoorOrLiftID { get; set; }

        public int ZoneID { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] CardNumber { get; set; }

        public int CardFacilityCode { get; set; }

        public int ServerID { get; set; }

        public int Sequence { get; set; }

        public int? EntrySequence { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
