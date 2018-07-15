namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class ClubAlarmZone
    {
        public int AccessGroupID { get; set; }

        public int AlarmZoneID { get; set; }

        public Guid ID { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
