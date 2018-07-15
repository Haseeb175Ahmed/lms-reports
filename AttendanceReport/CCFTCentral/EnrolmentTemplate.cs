namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EnrolmentTemplate")]
    public partial class EnrolmentTemplate
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public EnrolmentTemplate()
        {
            EnrolmentTemplateDatas = new HashSet<EnrolmentTemplateData>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? PrimaryPDFID { get; set; }

        public int? CardholderPDFID { get; set; }

        public int? OtherPDFID { get; set; }

        public int? CardTypeID { get; set; }

        public int CreateNewCard { get; set; }

        [StringLength(100)]
        public string DefaultReason { get; set; }

        public int PrintEncodeOption { get; set; }

        public int PrintOption { get; set; }

        public int EncodeOption { get; set; }

        public int UseCardTypeDates { get; set; }

        public virtual CardType CardType { get; set; }

        public virtual PersonalDataField PersonalDataField { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual PersonalDataField PersonalDataField1 { get; set; }

        public virtual PersonalDataField PersonalDataField2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EnrolmentTemplateData> EnrolmentTemplateDatas { get; set; }
    }
}
