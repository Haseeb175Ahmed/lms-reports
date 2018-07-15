namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Club")]
    public partial class Club
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Club()
        {
            AccessTriplets = new HashSet<AccessTriplet>();
            Club1 = new HashSet<Club>();
            ClubMembershipPairs = new HashSet<ClubMembershipPair>();
            ClubPersonalDataFieldPairs = new HashSet<ClubPersonalDataFieldPair>();
            SaltoAccessTriplets = new HashSet<SaltoAccessTriplet>();
            Club11 = new HashSet<Club>();
            Clubs = new HashSet<Club>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? ParentClubID { get; set; }

        public int PrivilegeFlags { get; set; }

        public DateTime? FromWhenDefault { get; set; }

        public DateTime? UntilWhenDefault { get; set; }

        public bool AutoRemoveExpired { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessTriplet> AccessTriplets { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Club> Club1 { get; set; }

        public virtual Club Club2 { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ClubMembershipPair> ClubMembershipPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ClubPersonalDataFieldPair> ClubPersonalDataFieldPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SaltoAccessTriplet> SaltoAccessTriplets { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Club> Club11 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Club> Clubs { get; set; }
    }
}
