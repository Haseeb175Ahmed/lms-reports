namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ChangeSetItem")]
    public partial class ChangeSetItem
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ChangeSetItem()
        {
            ChangeSetFields = new HashSet<ChangeSetField>();
        }

        public int ChangeSetItemID { get; set; }

        public Guid ChangeSetID { get; set; }

        public Guid ItemID { get; set; }

        public Guid ItemType { get; set; }

        public byte ChangeType { get; set; }

        public virtual ChangeSet ChangeSet { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ChangeSetField> ChangeSetFields { get; set; }
    }
}
