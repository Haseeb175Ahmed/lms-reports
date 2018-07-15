namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("PersonalDataStrList")]
    public partial class PersonalDataStrList
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int PersonalDataFieldID { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(200)]
        public string Value { get; set; }

        [Key]
        [Column(Order = 2)]
        public Guid GlobalID { get; set; }

        public virtual PersonalDataField PersonalDataField { get; set; }
    }
}
