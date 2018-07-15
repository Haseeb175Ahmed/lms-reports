namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("RelationshipInfo")]
    public partial class RelationshipInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Relationship { get; set; }

        public bool Symmetrical { get; set; }

        public bool DoPush { get; set; }

        public int PreventDeletionCode { get; set; }
    }
}
