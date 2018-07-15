namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorSession")]
    public partial class OperatorSession
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public OperatorSession()
        {
            SessionOperatorAccesses = new HashSet<SessionOperatorAccess>();
            SessionOperatorClubs = new HashSet<SessionOperatorClub>();
            SessionPermissions = new HashSet<SessionPermission>();
        }

        public int ID { get; set; }

        public int OperatorID { get; set; }

        public int? WorkstationID { get; set; }

        public DateTime LogonTime { get; set; }

        public int? CurrentDivisionID { get; set; }

        public int Orphaned { get; set; }

        public virtual Division Division { get; set; }

        public virtual Operator Operator { get; set; }

        public virtual Workstation Workstation { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SessionOperatorAccess> SessionOperatorAccesses { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SessionOperatorClub> SessionOperatorClubs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SessionPermission> SessionPermissions { get; set; }
    }
}
