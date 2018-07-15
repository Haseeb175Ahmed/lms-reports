namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class ExternalSystemType
    {
        [Key]
        [StringLength(64)]
        public string InternalName { get; set; }

        public int DisplayName { get; set; }

        [Required]
        [StringLength(256)]
        public string TypeID { get; set; }

        public Guid? ESPropertiesCLSID { get; set; }

        public Guid? ESIPropertiesCLSID { get; set; }
    }
}
