namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FenceZone")]
    public partial class FenceZone
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public bool DataLogFenceVoltageIntervalEnabled { get; set; }

        public bool DataLogFenceVoltageChangeEnabled { get; set; }

        public int DataLogFenceVoltageInterval { get; set; }

        public float DataLogFenceVoltageChange { get; set; }

        public float FenceWarningVoltageLowLevel { get; set; }

        public float FenceWarningVoltageHighLevel { get; set; }

        public int FenceWarningVoltagePulseCount { get; set; }

        public float FenceAlertVoltageLowLevel { get; set; }

        public float FenceAlertVoltageHighLevel { get; set; }

        public int FenceAlertVoltagePulseCount { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
