namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("EDIConfigData")]
    public partial class EDIConfigData
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ConfigSection { get; set; }

        [Key]
        [Column(Order = 2)]
        [StringLength(80)]
        public string ConfigKey { get; set; }

        public int? ConfigOrder { get; set; }

        [Key]
        [Column(Order = 3)]
        [StringLength(2048)]
        public string ConfigValue { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
