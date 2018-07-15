namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("PersonalDataField")]
    public partial class PersonalDataField
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public PersonalDataField()
        {
            ClubPersonalDataFieldPairs = new HashSet<ClubPersonalDataFieldPair>();
            EnrolmentTemplates = new HashSet<EnrolmentTemplate>();
            EnrolmentTemplates1 = new HashSet<EnrolmentTemplate>();
            EnrolmentTemplates2 = new HashSet<EnrolmentTemplate>();
            OperatorClubPDFAccessPairs = new HashSet<OperatorClubPDFAccessPair>();
            EnrolmentTemplateDatas = new HashSet<EnrolmentTemplateData>();
            PersonalDataImageIDs = new HashSet<PersonalDataImageID>();
            PersonalDataIntegers = new HashSet<PersonalDataInteger>();
            PersonalDataStrings = new HashSet<PersonalDataString>();
            PersonalDataStrLists = new HashSet<PersonalDataStrList>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int ValueTypeCode { get; set; }

        public int IsRequiredField { get; set; }

        public int IsUniqueField { get; set; }

        public int HasNoDefault { get; set; }

        public int DefaultAccess { get; set; }

        public int SortPriority { get; set; }

        [Required]
        [MaxLength(1000)]
        public byte[] ConfigurationData { get; set; }

        public bool RestrictedList { get; set; }

        public DateTime? DateDefaultValue { get; set; }

        public DateTime? DateLowerLimit { get; set; }

        public DateTime? DateUpperLimit { get; set; }

        public int UseRegularExpression { get; set; }

        [StringLength(2048)]
        public string RegularExpressionSyntax { get; set; }

        [StringLength(2048)]
        public string RegularExpressionDesc { get; set; }

        public int BaseValueTypeCode { get; set; }

        public bool NotificationCapable { get; set; }

        public bool NotificationsEnabled { get; set; }

        public Guid? NotificationGlobalID { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ClubPersonalDataFieldPair> ClubPersonalDataFieldPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EnrolmentTemplate> EnrolmentTemplates { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EnrolmentTemplate> EnrolmentTemplates1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EnrolmentTemplate> EnrolmentTemplates2 { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubPDFAccessPair> OperatorClubPDFAccessPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EnrolmentTemplateData> EnrolmentTemplateDatas { get; set; }

        public virtual SiteNotifyPDF SiteNotifyPDF { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalDataImageID> PersonalDataImageIDs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalDataInteger> PersonalDataIntegers { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalDataString> PersonalDataStrings { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PersonalDataStrList> PersonalDataStrLists { get; set; }
    }
}
