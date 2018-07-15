namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SitePlan")]
    public partial class SitePlan
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public SitePlan()
        {
            SitePlan1 = new HashSet<SitePlan>();
            SitePlanGraphics = new HashSet<SitePlanGraphic>();
            SitePlanLinks = new HashSet<SitePlanLink>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Column(TypeName = "image")]
        public byte[] Picture { get; set; }

        public int PictureSize { get; set; }

        public int Format { get; set; }

        public int? ParentSitePlanID { get; set; }

        public int? IconSetID { get; set; }

        public int? IconScaleFactorNum { get; set; }

        public int OriginalWindowWidth { get; set; }

        public int OriginalWindowHeight { get; set; }

        public int? IconScaleFactorDen { get; set; }

        public int BorderColour { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual IconSet IconSet { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SitePlan> SitePlan1 { get; set; }

        public virtual SitePlan SitePlan2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SitePlanGraphic> SitePlanGraphics { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SitePlanLink> SitePlanLinks { get; set; }
    }
}
