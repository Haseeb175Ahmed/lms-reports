namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DivisionFindVisitorGridColumn")]
    public partial class DivisionFindVisitorGridColumn
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int DisplayOrder { get; set; }

        [Key]
        [Column(Order = 1)]
        public int ID { get; set; }

        public int FieldID { get; set; }

        public int? PDFID { get; set; }

        public bool IncludeInSearch { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }
    }
}
