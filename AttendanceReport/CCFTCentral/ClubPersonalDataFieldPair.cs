namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ClubPersonalDataFieldPair")]
    public partial class ClubPersonalDataFieldPair
    {
        public int ClubID { get; set; }

        public int PersonalDataFieldID { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual Club Club { get; set; }

        public virtual PersonalDataField PersonalDataField { get; set; }
    }
}
