namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ExternalEventSource")]
    public partial class ExternalEventSource
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [StringLength(64)]
        public string SourceID { get; set; }

        public int DiallingXYZ { get; set; }

        public int DiallingGroup { get; set; }

        [Required]
        [StringLength(100)]
        public string EventTriggerString { get; set; }

        [Column(TypeName = "ntext")]
        [Required]
        public string ItemConfig { get; set; }

        [Required]
        [StringLength(64)]
        public string ItemConfigType { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }
    }
}
