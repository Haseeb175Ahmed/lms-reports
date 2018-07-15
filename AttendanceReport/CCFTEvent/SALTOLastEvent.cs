namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SALTOLastEvent")]
    public partial class SALTOLastEvent
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SaltoServerID { get; set; }

        [StringLength(500)]
        public string RowString { get; set; }
    }
}
