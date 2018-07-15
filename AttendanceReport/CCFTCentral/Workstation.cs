namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Workstation")]
    public partial class Workstation
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Workstation()
        {
            OperatorSessions = new HashSet<OperatorSession>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int IsRestricted { get; set; }

        public int AutoLogoff { get; set; }

        public int LogonSchemes { get; set; }

        public int CardRemovalTime { get; set; }

        public int CardWarningTime { get; set; }

        public bool DisableExit { get; set; }

        public int? VMSDoorID { get; set; }

        public int? VMSReceptionID { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        public virtual FTItem FTItem2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorSession> OperatorSessions { get; set; }
    }
}
