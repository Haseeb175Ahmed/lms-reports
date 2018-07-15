namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MiscellaneousIcon")]
    public partial class MiscellaneousIcon
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int MiscIconIndex { get; set; }

        [Required]
        public byte[] IconImage { get; set; }
    }
}
