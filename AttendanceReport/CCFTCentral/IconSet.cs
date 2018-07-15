namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("IconSet")]
    public partial class IconSet
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public IconSet()
        {
            FTItemTypes = new HashSet<FTItemType>();
            IconSetStateSettings = new HashSet<IconSetStateSetting>();
            SiteItems = new HashSet<SiteItem>();
            SitePlans = new HashSet<SitePlan>();
            SitePlanLinks = new HashSet<SitePlanLink>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int FTItemTypeID { get; set; }

        public int? isSystemIconSet { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItemType> FTItemTypes { get; set; }

        public virtual FTItemType FTItemType { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<IconSetStateSetting> IconSetStateSettings { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SiteItem> SiteItems { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SitePlan> SitePlans { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SitePlanLink> SitePlanLinks { get; set; }
    }
}
