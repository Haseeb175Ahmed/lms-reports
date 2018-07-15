namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CmdrOutput")]
    public partial class CmdrOutput
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? LogicalAddress { get; set; }

        public int? PulseTime { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(100)]
        public string OnStateName { get; set; }

        [Key]
        [Column(Order = 2)]
        [StringLength(100)]
        public string OffStateName { get; set; }
    }
}
