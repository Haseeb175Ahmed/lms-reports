namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Lookup")]
    public partial class Lookup
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long EventID { get; set; }
    }
}
