namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmPriority")]
    public partial class AlarmPriority
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Priority { get; set; }

        [Required]
        [StringLength(50)]
        public string Name { get; set; }

        public int Colour { get; set; }

        public int? SoundID { get; set; }

        public int Tolerance { get; set; }

        public int ForgetFloodInterval { get; set; }

        public bool SendMultibreaks { get; set; }

        [StringLength(260)]
        public string AudioFile { get; set; }

        [Column(TypeName = "image")]
        public byte[] AudioData { get; set; }
    }
}
