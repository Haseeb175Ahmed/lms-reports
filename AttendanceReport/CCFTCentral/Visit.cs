namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Visit")]
    public partial class Visit
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? PlannedReception { get; set; }

        public int? RegisteredBy { get; set; }

        public DateTime RegistrationTime { get; set; }

        public int? VisitorType { get; set; }

        public int Status { get; set; }

        [Required]
        [StringLength(200)]
        public string BadgeText { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }

        public virtual FTItem FTItem2 { get; set; }

        public virtual FTItem FTItem3 { get; set; }
    }
}
