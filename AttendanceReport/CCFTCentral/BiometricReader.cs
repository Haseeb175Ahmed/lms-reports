namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("BiometricReader")]
    public partial class BiometricReader
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int IPAddress { get; set; }

        public int Port { get; set; }

        public int FalseAcceptanceRate { get; set; }

        public int AuthenticationMode { get; set; }

        public bool AutoSynchronisation { get; set; }

        public bool PollEnabled { get; set; }

        public int PollTime { get; set; }

        public bool DisplayNames { get; set; }

        public bool AlwaysRequestPIN { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
