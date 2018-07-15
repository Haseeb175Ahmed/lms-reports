namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorClubPDFAccessPair")]
    public partial class OperatorClubPDFAccessPair
    {
        public int OperatorClubID { get; set; }

        public int PersonalDataFieldID { get; set; }

        public int AccessLevel { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual OperatorClub OperatorClub { get; set; }

        public virtual PersonalDataField PersonalDataField { get; set; }
    }
}
