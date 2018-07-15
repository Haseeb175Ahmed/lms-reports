namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FenceController")]
    public partial class FenceController
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public float PulseInterval { get; set; }

        public int PowerLevel { get; set; }

        public int SynchronisationSource { get; set; }

        public float LowBatteryLevel { get; set; }

        public float GoSlowPulseInterval { get; set; }

        public int GoSlowOnMainsFail { get; set; }

        public int GoSlowMainsFailDelay { get; set; }

        public int GoSlowPowerLevel { get; set; }

        public float GoSlowBatteryLevel { get; set; }

        public float StopBatteryLevel { get; set; }

        public bool DataLogEarthVoltageIntervalEnabled { get; set; }

        public bool DataLogEarthVoltageChangeEnabled { get; set; }

        public bool DataLogBatteryVoltageIntervalEnabled { get; set; }

        public bool DataLogBatteryVoltageChangeEnabled { get; set; }

        public int DataLogEarthVoltageInterval { get; set; }

        public float DataLogEarthVoltageChange { get; set; }

        public int DataLogBatteryVoltageInterval { get; set; }

        public float DataLogBatteryVoltageChange { get; set; }

        public float EarthAlertVoltageLowLevel { get; set; }

        public float EarthAlertVoltageHighLevel { get; set; }

        public int EarthAlertPulseCount { get; set; }

        public int RelayExpanderExists { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
