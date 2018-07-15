namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SADPlugin")]
    public partial class SADPlugin
    {
        [Key]
        public Guid SadPluginClassID { get; set; }

        [Required]
        [StringLength(64)]
        public string SadPluginName { get; set; }

        public int SadPluginSet { get; set; }

        public int InstantiationOrder { get; set; }
    }
}
