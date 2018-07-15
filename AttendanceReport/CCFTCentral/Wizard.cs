namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Wizard
    {
        public int WizardType { get; set; }

        [Key]
        public Guid ClassID { get; set; }

        public int BusRequirements { get; set; }

        public Guid TemplateClassID { get; set; }

        public int TitleID { get; set; }

        [Required]
        [StringLength(128)]
        public string Description { get; set; }
    }
}
