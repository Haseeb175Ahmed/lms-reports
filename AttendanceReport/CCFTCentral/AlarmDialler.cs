namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmDialler")]
    public partial class AlarmDialler
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int RedialAttempts { get; set; }

        public int RedialInterval { get; set; }

        public bool UseNumber1 { get; set; }

        [Required]
        [StringLength(33)]
        public string Number1 { get; set; }

        [Required]
        [StringLength(33)]
        public string Description1 { get; set; }

        public bool UseNumber2 { get; set; }

        [Required]
        [StringLength(33)]
        public string Number2 { get; set; }

        [Required]
        [StringLength(33)]
        public string Description2 { get; set; }

        public bool UseNumber3 { get; set; }

        [Required]
        [StringLength(33)]
        public string Number3 { get; set; }

        [Required]
        [StringLength(33)]
        public string Description3 { get; set; }

        public bool UsePrefix1 { get; set; }

        [Required]
        [StringLength(33)]
        public string Prefix1 { get; set; }

        public bool DoPDTC { get; set; }

        public int PDTCInterval { get; set; }

        public int PERDRhythm { get; set; }

        public DateTime PERDTimeOfDay { get; set; }

        public int PERDInterval { get; set; }

        public DateTime PERDStart { get; set; }

        public bool PERDSkip { get; set; }

        public int MT { get; set; }

        public int HSWait { get; set; }

        public int TxAttempts { get; set; }

        [Required]
        [StringLength(33)]
        public string ModemInit { get; set; }

        public int ModemChipset { get; set; }

        public int Backup1 { get; set; }

        public int Backup2 { get; set; }

        public int Backup3 { get; set; }

        public int Backup4 { get; set; }

        public int Backup5 { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
