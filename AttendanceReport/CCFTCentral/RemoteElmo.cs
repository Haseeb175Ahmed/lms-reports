namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("RemoteElmo")]
    public partial class RemoteElmo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Required]
        [StringLength(255)]
        public string HostName { get; set; }

        public int Port { get; set; }

        public int ResponseTimeout { get; set; }

        public int OfflineTimeout { get; set; }

        public int PublishInterval { get; set; }

        [Required]
        [MaxLength(8000)]
        public byte[] Certificate { get; set; }

        public bool CommunicationEnabled { get; set; }

        [Required]
        [StringLength(30)]
        public string Version { get; set; }

        public int LicenceNumber { get; set; }

        public int RemainingChangeSets { get; set; }

        public int CardholderLocationSequence { get; set; }

        public bool Aggregation { get; set; }

        public Guid ChangeSetID { get; set; }

        public int ChangeSetSequence { get; set; }

        public Guid LocalChangeSetID { get; set; }

        public int LocalChangeSetSequence { get; set; }

        public int? AlarmZoneSetStringID { get; set; }

        public int? AlarmZoneUnsetStringID { get; set; }

        [StringLength(255)]
        public string AlarmZoneUser1StateName { get; set; }

        [StringLength(255)]
        public string AlarmZoneUser2StateName { get; set; }

        public virtual SiteItem SiteItem { get; set; }
    }
}
