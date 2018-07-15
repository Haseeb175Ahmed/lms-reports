namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("PersonalDataInteger")]
    public partial class PersonalDataInteger
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int PersonalDataFieldID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CardholderID { get; set; }

        public int? Value { get; set; }

        public virtual Cardholder Cardholder { get; set; }

        public virtual PersonalDataField PersonalDataField { get; set; }
    }
}
