namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("PromptEntry")]
    public partial class PromptEntry
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int PromptID { get; set; }

        public int? Row { get; set; }

        public int? Col { get; set; }

        public int? Type { get; set; }

        public int? Parameter { get; set; }

        [StringLength(127)]
        public string PromptString { get; set; }
    }
}
