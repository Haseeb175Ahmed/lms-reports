namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SaltoServer")]
    public partial class SaltoServer
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Required]
        [StringLength(255)]
        public string HostName { get; set; }

        public int IPAddress { get; set; }

        public int OutgoingPort { get; set; }

        public int IncomingPort { get; set; }

        public int HoursOfAccess { get; set; }

        public int HoursOf247Access { get; set; }

        public bool DefaultOfficePrivilege { get; set; }

        public bool DefaultLowBatteryWarn { get; set; }

        public bool DefaultOverridePrivacy { get; set; }

        public bool DefaultAccommodationMode { get; set; }

        public bool DefaultESDRelay1 { get; set; }

        public bool DefaultESDRelay2 { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
