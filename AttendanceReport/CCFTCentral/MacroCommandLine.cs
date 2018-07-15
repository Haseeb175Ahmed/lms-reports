namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MacroCommandLine")]
    public partial class MacroCommandLine
    {
        [Key]
        public int CommandLineID { get; set; }

        public int FTItemID { get; set; }

        [Required]
        [StringLength(260)]
        public string CommandLine { get; set; }

        [Required]
        [StringLength(300)]
        public string CommandArgument { get; set; }

        public virtual Macro Macro { get; set; }
    }
}
