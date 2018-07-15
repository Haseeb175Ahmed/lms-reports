namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SaltoAccessTriplet")]
    public partial class SaltoAccessTriplet
    {
        public int ClubID { get; set; }

        public int SaltoAccessZoneID { get; set; }

        public int ScheduleID { get; set; }

        public Guid ID { get; set; }

        public virtual Club Club { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
