namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardType")]
    public partial class CardType
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public CardType()
        {
            AperioDoors = new HashSet<AperioDoor>();
            Cards = new HashSet<Card>();
            EnrolmentTemplates = new HashSet<EnrolmentTemplate>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int FacilityCode { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] MinimumNumber { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] MaximumNumber { get; set; }

        public DateTime? ActivationDate { get; set; }

        public DateTime? ExpiryDate { get; set; }

        public int IssueLevel { get; set; }

        public int InactivityTimeout { get; set; }

        public int PinLength { get; set; }

        public bool UseForCardholder { get; set; }

        public int? LastEncodedNumber { get; set; }

        public bool SaltoEventsOnCards { get; set; }

        [StringLength(32)]
        public string MifareSiteKey { get; set; }

        public Guid CardNumberFormatClassID { get; set; }

        public int WarningPeriod { get; set; }

        public int WarningPeriodType { get; set; }

        public bool EnableExpiryNotify { get; set; }

        [Required]
        [StringLength(32)]
        public string MifarePlusSiteKey { get; set; }

        public bool EncodeAtMifarePlusSL3 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AperioDoor> AperioDoors { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Card> Cards { get; set; }

        public virtual CardLayout CardLayout { get; set; }

        public virtual CardNumberFormat CardNumberFormat { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EnrolmentTemplate> EnrolmentTemplates { get; set; }
    }
}
