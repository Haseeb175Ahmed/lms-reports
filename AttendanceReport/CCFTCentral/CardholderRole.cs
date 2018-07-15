namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class CardholderRole
    {
        public int CardholderID { get; set; }

        public int RoleID { get; set; }

        public int RoleCardholderID { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual Cardholder Cardholder { get; set; }

        public virtual Cardholder Cardholder1 { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
