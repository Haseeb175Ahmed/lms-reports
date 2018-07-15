namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("LiftHLI")]
    public partial class LiftHLI
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int OfflineTimeout { get; set; }

        [Required]
        [StringLength(1024)]
        public string ConfigurationString { get; set; }

        public int Protocol { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }
    }
}
