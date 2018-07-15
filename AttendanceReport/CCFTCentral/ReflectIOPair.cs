namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ReflectIOPair")]
    public partial class ReflectIOPair
    {
        public int? OutputID { get; set; }

        public int? InputID { get; set; }

        [Key]
        public bool Inverted { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
