namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class CardholderNotificationFilter
    {
        public int CardholderID { get; set; }

        public int FilterID { get; set; }

        public int ScheduleID { get; set; }

        public int FilterEnabled { get; set; }

        public int NotifyMethod { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual Cardholder Cardholder { get; set; }
    }
}
