namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CmdrInput")]
    public partial class CmdrInput
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(100)]
        public string OpenMessage { get; set; }

        [Key]
        [Column(Order = 2)]
        [StringLength(100)]
        public string CloseMessage { get; set; }

        [Key]
        [Column(Order = 3)]
        [StringLength(100)]
        public string TamperMessage { get; set; }

        [Key]
        [Column(Order = 4)]
        [StringLength(100)]
        public string OpenStateName { get; set; }

        [Key]
        [Column(Order = 5)]
        [StringLength(100)]
        public string CloseStateName { get; set; }

        [Key]
        [Column(Order = 6)]
        [StringLength(100)]
        public string TamperStateName { get; set; }
    }
}
