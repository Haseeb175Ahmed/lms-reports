namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MacroAction")]
    public partial class MacroAction
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ActionItemID { get; set; }

        public int VerbID { get; set; }

        public int Duration { get; set; }

        [Required]
        [MaxLength(4096)]
        public byte[] ActionData { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual Macro Macro { get; set; }
    }
}
