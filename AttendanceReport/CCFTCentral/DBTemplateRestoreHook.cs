namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DBTemplateRestoreHook")]
    public partial class DBTemplateRestoreHook
    {
        [Key]
        [Column(Order = 0)]
        public Guid ClassID { get; set; }

        [Key]
        [Column(Order = 1)]
        public string HookName { get; set; }
    }
}
