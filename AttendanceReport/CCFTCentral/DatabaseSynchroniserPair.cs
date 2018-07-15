namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DatabaseSynchroniserPair")]
    public partial class DatabaseSynchroniserPair
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int DatabaseType { get; set; }

        public Guid DatabaseSynchroniser { get; set; }
    }
}
