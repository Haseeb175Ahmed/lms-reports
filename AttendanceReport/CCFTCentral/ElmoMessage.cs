namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ElmoMessage")]
    public partial class ElmoMessage
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int QueueID { get; set; }

        [Required]
        [MaxLength(1000)]
        public byte[] Body { get; set; }

        public int Priority { get; set; }

        public int MaxSecondsQueued { get; set; }

        public DateTime TimeAddedToQueue { get; set; }

        public int State { get; set; }

        public int BodyChecksum { get; set; }
    }
}
