namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorClubPrivilegePair")]
    public partial class OperatorClubPrivilegePair
    {
        public int OperatorClubID { get; set; }

        public int PrivilegeID { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual OperatorClub OperatorClub { get; set; }

        public virtual OperatorPrivilege OperatorPrivilege { get; set; }
    }
}
