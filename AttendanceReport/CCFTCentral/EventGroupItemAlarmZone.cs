namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EventGroupItemAlarmZone")]
    public partial class EventGroupItemAlarmZone
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SiteItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int EventGroupID { get; set; }

        public int AlarmZoneID { get; set; }

        public virtual AlarmZone AlarmZone { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
