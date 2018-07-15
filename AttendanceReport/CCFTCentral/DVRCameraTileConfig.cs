namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DVRCameraTileConfig")]
    public partial class DVRCameraTileConfig
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid GlobalID { get; set; }

        public short CameraConfigType { get; set; }

        public int SpecificCameraItemID { get; set; }

        [Required]
        [StringLength(200)]
        public string SpecificCameraPreset { get; set; }

        public bool SpecificCameraAllowPtz { get; set; }

        public short EventCamera { get; set; }

        public short EventCameraNumber { get; set; }

        public bool EventCameraAlwaysDisplayLive { get; set; }

        public bool EventCameraAllowPtzForLive { get; set; }

        [Key]
        [Column(Order = 2)]
        public Guid PanelGlobalID { get; set; }
    }
}
