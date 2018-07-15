namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("VisitVisitor")]
    public partial class VisitVisitor
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public VisitVisitor()
        {
            VisitVisitorAccessGroups = new HashSet<VisitVisitorAccessGroup>();
            VisitVisitorActions = new HashSet<VisitVisitorAction>();
        }

        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CardholderID { get; set; }

        public int Status { get; set; }

        [Required]
        [MaxLength(256)]
        public byte[] EncodedNumber { get; set; }

        public int? CardTypeID { get; set; }

        public DateTime? ActualOnSiteTime { get; set; }

        public DateTime? ActualLeaveSiteTime { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        public virtual FTItem FTItem2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitorAccessGroup> VisitVisitorAccessGroups { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitorAction> VisitVisitorActions { get; set; }
    }
}
