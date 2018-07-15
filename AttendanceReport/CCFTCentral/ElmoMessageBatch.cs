namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ElmoMessageBatch")]
    public partial class ElmoMessageBatch
    {
        [Key]
        [Column(Order = 0)]
        public int ID { get; set; }

        public int QueueID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Priority { get; set; }

        public byte[] MessageData { get; set; }
    }
}
