namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ElmoPersonalDataFieldPair")]
    public partial class ElmoPersonalDataFieldPair
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int PersonalDataFieldID { get; set; }

        public int FontStyle { get; set; }

        public int SortOrder { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
