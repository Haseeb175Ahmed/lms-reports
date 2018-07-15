namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CompetencyMessageInfo")]
    public partial class CompetencyMessageInfo
    {
        public int? FTItemID { get; set; }

        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int TypeID { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(2048)]
        public string TheMessage { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int TheMessageDelay { get; set; }

        public virtual Competency Competency { get; set; }
    }
}
