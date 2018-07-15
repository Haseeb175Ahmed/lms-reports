namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ChangeSetField")]
    public partial class ChangeSetField
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ChangeSetItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid PropertyID { get; set; }

        public int TimeOverride { get; set; }

        public virtual ChangeSetItem ChangeSetItem { get; set; }
    }
}
