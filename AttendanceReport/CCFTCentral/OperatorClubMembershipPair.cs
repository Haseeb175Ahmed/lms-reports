namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorClubMembershipPair")]
    public partial class OperatorClubMembershipPair
    {
        public int OperatorClubID { get; set; }

        public int OperatorID { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual Cardholder Cardholder { get; set; }

        public virtual OperatorClub OperatorClub { get; set; }
    }
}
