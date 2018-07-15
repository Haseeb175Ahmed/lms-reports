namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Cardholder")]
    public partial class Cardholder
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Cardholder()
        {
            Cards = new HashSet<Card>();
            CardholderCompetencyPairs = new HashSet<CardholderCompetencyPair>();
            CardholderFingerprints = new HashSet<CardholderFingerprint>();
            CardholderNotificationFilters = new HashSet<CardholderNotificationFilter>();
            CardholderRoles = new HashSet<CardholderRole>();
            CardholderRoles1 = new HashSet<CardholderRole>();
            ClubMembershipPairs = new HashSet<ClubMembershipPair>();
            OperatorClubMembershipPairs = new HashSet<OperatorClubMembershipPair>();
            OperatorExpiredPasswords = new HashSet<OperatorExpiredPassword>();
            OperatorPreferences = new HashSet<OperatorPreference>();
            PersonalDataImageIDs = new HashSet<PersonalDataImageID>();
            PersonalDataIntegers = new HashSet<PersonalDataInteger>();
            PersonalDataStrings = new HashSet<PersonalDataString>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Required]
        [StringLength(50)]
        public string FirstName { get; set; }

        [Required]
        [StringLength(50)]
        public string LastName { get; set; }

        public int Authorised { get; set; }

        [Required]
        [StringLength(4)]
        public string DiallingCCC { get; set; }

        [MaxLength(28)]
        public byte[] EncryptedUserCode { get; set; }

        public int UseAlternateUnlockTime { get; set; }

        public int UserCodeConflictCount { get; set; }

        public int? VisitorResponsibilityID { get; set; }

        public int NotificationsEnabled { get; set; }

        public DateTime? NotifyFrom { get; set; }

        public DateTime? NotifyUntil { get; set; }

        public bool? SaltoOfficePrivilege { get; set; }

        public bool? SaltoWarnLowBattery { get; set; }

        public bool? SaltoOverridePrivacy { get; set; }

        public bool? SaltoAccommodationMode { get; set; }

        public bool? SaltoESDRelay1 { get; set; }

        public bool? SaltoESDRelay2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Card> Cards { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderCompetencyPair> CardholderCompetencyPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderFingerprint> CardholderFingerprints { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderNotificationFilter> CardholderNotificationFilters { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderRole> CardholderRoles { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderRole> CardholderRoles1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ClubMembershipPair> ClubMembershipPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubMembershipPair> OperatorClubMembershipPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorExpiredPassword> OperatorExpiredPasswords { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorPreference> OperatorPreferences { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalDataImageID> PersonalDataImageIDs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalDataInteger> PersonalDataIntegers { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalDataString> PersonalDataStrings { get; set; }
    }
}
