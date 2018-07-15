namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("WizardTemplate")]
    public partial class WizardTemplate
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Column(TypeName = "ntext")]
        public string TemplateXMLData { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
