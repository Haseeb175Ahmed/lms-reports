namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmInstructionTag")]
    public partial class AlarmInstructionTag
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AlarmInstructionID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int InsertionTagID { get; set; }

        public int cReference { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual InsertionTag InsertionTag { get; set; }
    }
}
