namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DoorPersonalDataFieldPair")]
    public partial class DoorPersonalDataFieldPair
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int DoorID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int PersonalDataFieldID { get; set; }

        public int FontStyle { get; set; }

        public int SortOrder { get; set; }

        public virtual Door Door { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
