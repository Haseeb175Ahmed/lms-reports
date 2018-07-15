namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SharedBinary")]
    public partial class SharedBinary
    {
        public Guid ID { get; set; }

        [Required]
        public byte[] BinaryData { get; set; }

        public int Checksum { get; set; }
    }
}
