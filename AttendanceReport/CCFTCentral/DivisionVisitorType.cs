namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DivisionVisitorType")]
    public partial class DivisionVisitorType
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public DivisionVisitorType()
        {
            DivisionVisitorTypeAccessGroups = new HashSet<DivisionVisitorTypeAccessGroup>();
            DivisionVisitorTypeFields = new HashSet<DivisionVisitorTypeField>();
            DivisionVisitorTypeHostTypes = new HashSet<DivisionVisitorTypeHostType>();
            DivisionVisitorTypeVisitVisitorFields = new HashSet<DivisionVisitorTypeVisitVisitorField>();
        }

        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int VisitorTypeID { get; set; }

        public int? CardTypeID { get; set; }

        public int VisitorNameOption { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        public virtual FTItem FTItem2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeAccessGroup> DivisionVisitorTypeAccessGroups { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeField> DivisionVisitorTypeFields { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeHostType> DivisionVisitorTypeHostTypes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeVisitVisitorField> DivisionVisitorTypeVisitVisitorFields { get; set; }
    }
}
