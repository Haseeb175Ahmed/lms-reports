namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SitePlanLink")]
    public partial class SitePlanLink
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public SitePlanLink()
        {
            SitePlanGraphics = new HashSet<SitePlanGraphic>();
        }

        public int ID { get; set; }

        public int? IconSetID { get; set; }

        public int TargetSitePlanID { get; set; }

        public virtual IconSet IconSet { get; set; }

        public virtual SitePlan SitePlan { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SitePlanGraphic> SitePlanGraphics { get; set; }
    }
}
