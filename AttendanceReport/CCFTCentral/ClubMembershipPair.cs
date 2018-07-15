namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ClubMembershipPair")]
    public partial class ClubMembershipPair
    {
        public int ID { get; set; }

        public int ClubID { get; set; }

        public int CardholderID { get; set; }

        public DateTime? FromWhen { get; set; }

        public DateTime? UntilWhen { get; set; }

        public Guid GlobalID { get; set; }

        public virtual Cardholder Cardholder { get; set; }

        public virtual Club Club { get; set; }
    }
}
