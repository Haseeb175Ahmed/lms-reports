namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ChangeSet")]
    public partial class ChangeSet
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ChangeSet()
        {
            ChangeSetItems = new HashSet<ChangeSetItem>();
        }

        public Guid ChangeSetID { get; set; }

        public Guid ServerID { get; set; }

        public int Sequence { get; set; }

        public DateTime Time { get; set; }

        public Guid FTItemID { get; set; }

        public Guid OperatorID { get; set; }

        public Guid WorkstationID { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ChangeSetItem> ChangeSetItems { get; set; }

        public virtual ChangeSetPending ChangeSetPending { get; set; }
    }
}
