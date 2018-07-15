namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Competency")]
    public partial class Competency
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Competency()
        {
            CardholderCompetencyPairs = new HashSet<CardholderCompetencyPair>();
            AccessZoneCompetencyPairs = new HashSet<AccessZoneCompetencyPair>();
            CompetencyMessageInfoes = new HashSet<CompetencyMessageInfo>();
            OperatorClubCompetencyAccessPairs = new HashSet<OperatorClubCompetencyAccessPair>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int WarningPeriod { get; set; }

        public int WarningPeriodType { get; set; }

        public int DefaultAccess { get; set; }

        public int DefaultExpiryType { get; set; }

        public int DefaultExpiry { get; set; }

        public int? RecordNumber { get; set; }

        public bool EnableExpiryNotify { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderCompetencyPair> CardholderCompetencyPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessZoneCompetencyPair> AccessZoneCompetencyPairs { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CompetencyMessageInfo> CompetencyMessageInfoes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubCompetencyAccessPair> OperatorClubCompetencyAccessPairs { get; set; }
    }
}
