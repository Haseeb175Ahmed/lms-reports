namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ChangeSetPending")]
    public partial class ChangeSetPending
    {
        [Key]
        public Guid ChangeSetID { get; set; }

        public Guid ServerID { get; set; }

        public int Sequence { get; set; }

        public DateTime ArrivalTime { get; set; }

        public DateTime? ApplyTime { get; set; }

        [StringLength(200)]
        public string Result { get; set; }

        public bool Applied { get; set; }

        public virtual ChangeSet ChangeSet { get; set; }
    }
}
