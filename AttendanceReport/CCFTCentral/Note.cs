namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Note
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Column("Note")]
        [Required]
        [StringLength(2500)]
        public string Note1 { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
