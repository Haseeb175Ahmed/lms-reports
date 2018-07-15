namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DivisionVisitorTypeVisitVisitorField")]
    public partial class DivisionVisitorTypeVisitVisitorField
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int VisitorTypeID { get; set; }

        public int DisplayOrder { get; set; }

        [Key]
        [Column(Order = 2)]
        public int ID { get; set; }

        public int FieldID { get; set; }

        public int? PDFID { get; set; }

        public virtual DivisionVisitorType DivisionVisitorType { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        public virtual FTItem FTItem2 { get; set; }
    }
}
