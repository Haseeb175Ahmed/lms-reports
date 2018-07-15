namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class BulkChangeMod
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int BulkChangeID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTFieldID { get; set; }

        public int? IntParam1 { get; set; }

        public int? IntParam2 { get; set; }

        public int? IntParam3 { get; set; }

        public int? IntParam4 { get; set; }

        [StringLength(1000)]
        public string StrParam1 { get; set; }

        [StringLength(1000)]
        public string StrParam2 { get; set; }

        public DateTime? DateParam1 { get; set; }

        public DateTime? DateParam2 { get; set; }

        public int? IntParam5 { get; set; }

        public int? IntParam6 { get; set; }

        public virtual BulkChange BulkChange { get; set; }
    }
}
