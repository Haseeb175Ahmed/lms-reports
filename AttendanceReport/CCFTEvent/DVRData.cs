namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DVRData")]
    public partial class DVRData
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SequenceID { get; set; }

        [Required]
        [StringLength(110)]
        public string CameraIdentifier { get; set; }

        public Guid GUID { get; set; }

        public int PreSeconds { get; set; }

        public int PostSeconds { get; set; }

        public virtual Sequence Sequence { get; set; }
    }
}
