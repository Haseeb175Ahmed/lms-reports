using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("SystemSetting")]
    public class SystemSetting
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        
        public int DaysToEmailNotification { get; set; }

        public int DaysToBlockUser { get; set; }

        public string SmtpServer { get; set; }

        public string SmtpPort { get; set; }

        public string FromEmailAddress { get; set; }

        public string FromEmailPassword { get; set; }

        public bool IsSmptSSL { get; set; }

        public bool IsSmptAuthRequired { get; set; }
    }
}
