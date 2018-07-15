namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardEvent")]
    public partial class CardEvent
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long EventID { get; set; }

        public byte CardType { get; set; }

        [Required]
        [StringLength(1024)]
        public string CardNumber { get; set; }

        public int FacilityCode { get; set; }

        public byte IssueLevel { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] EncodedBinary { get; set; }

        public virtual Event Event { get; set; }
    }
}
