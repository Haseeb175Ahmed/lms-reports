namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("PanelListPanel")]
    public partial class PanelListPanel
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid GlobalID { get; set; }

        public int DisplayOrder { get; set; }

        public virtual ViewerPanel ViewerPanel { get; set; }
    }
}
