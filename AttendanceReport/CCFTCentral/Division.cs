namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Division")]
    public partial class Division
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Division()
        {
            Division1 = new HashSet<Division>();
            DivisionDescendants = new HashSet<DivisionDescendant>();
            DivisionDescendants1 = new HashSet<DivisionDescendant>();
            OperatorClubDivisions = new HashSet<OperatorClubDivision>();
            OperatorSessions = new HashSet<OperatorSession>();
            SessionOperatorAccesses = new HashSet<SessionOperatorAccess>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? ParentID { get; set; }

        public int VisitorOverdueGraceTime { get; set; }

        public bool ConfigurationActive { get; set; }

        [Required]
        [StringLength(32)]
        public string TimeZone { get; set; }

        public int StandardVisitStartTime { get; set; }

        public int StandardVisitEndTime { get; set; }

        public int HomeTimeRangePast { get; set; }

        public int HomeTimeRangeFuture { get; set; }

        public int? CardTypeID { get; set; }

        public int? VisitorDivisionID { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Division> Division1 { get; set; }

        public virtual Division Division2 { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        public virtual FTItem FTItem2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionDescendant> DivisionDescendants { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionDescendant> DivisionDescendants1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubDivision> OperatorClubDivisions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorSession> OperatorSessions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SessionOperatorAccess> SessionOperatorAccesses { get; set; }
    }
}
