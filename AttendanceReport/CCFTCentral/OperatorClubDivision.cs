namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorClubDivision")]
    public partial class OperatorClubDivision
    {
        public int OperatorClubID { get; set; }

        public int DivisionID { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual Division Division { get; set; }

        public virtual OperatorClub OperatorClub { get; set; }
    }
}
