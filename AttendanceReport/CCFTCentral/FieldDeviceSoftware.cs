namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FieldDeviceSoftware")]
    public partial class FieldDeviceSoftware
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FieldDeviceID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SoftwareType { get; set; }

        public int? DeviceSoftwareID { get; set; }

        [StringLength(12)]
        public string Version { get; set; }

        public virtual DeviceSoftware DeviceSoftware { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }
    }
}
