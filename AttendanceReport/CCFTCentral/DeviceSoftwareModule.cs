namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DeviceSoftwareModule")]
    public partial class DeviceSoftwareModule
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int DeviceSoftwareID { get; set; }

        public int ModuleType { get; set; }

        [Required]
        [StringLength(12)]
        public string ModuleFileName { get; set; }

        [Column(TypeName = "image")]
        public byte[] ModuleData { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ModuleIndex { get; set; }

        public virtual DeviceSoftware DeviceSoftware { get; set; }
    }
}
