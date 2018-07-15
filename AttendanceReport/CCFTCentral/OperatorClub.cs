namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorClub")]
    public partial class OperatorClub
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public OperatorClub()
        {
            OperatorClubCompetencyAccessPairs = new HashSet<OperatorClubCompetencyAccessPair>();
            OperatorClubDivisions = new HashSet<OperatorClubDivision>();
            OperatorClubMembershipPairs = new HashSet<OperatorClubMembershipPair>();
            OperatorClubPDFAccessPairs = new HashSet<OperatorClubPDFAccessPair>();
            OperatorClubPrivilegePairs = new HashSet<OperatorClubPrivilegePair>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int AlarmWindowAutoPopup { get; set; }

        public int AlarmWindowAlwaysOnTop { get; set; }

        public int AlarmWindowEnableEscalatedWindow { get; set; }

        public int AlarmWindowLeaveEscalatedAlarmsInMain { get; set; }

        public int AlarmWindowDenyAckProcessEscalated { get; set; }

        public int AlarmNotesOption { get; set; }

        public bool isRestricted { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubCompetencyAccessPair> OperatorClubCompetencyAccessPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubDivision> OperatorClubDivisions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubMembershipPair> OperatorClubMembershipPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubPDFAccessPair> OperatorClubPDFAccessPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubPrivilegePair> OperatorClubPrivilegePairs { get; set; }
    }
}
