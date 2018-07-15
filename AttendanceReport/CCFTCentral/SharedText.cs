namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SharedText")]
    public partial class SharedText
    {
        public Guid ID { get; set; }

        [Required]
        public string TextData { get; set; }

        public int Checksum { get; set; }
    }
}
