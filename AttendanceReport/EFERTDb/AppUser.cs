using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace AttendanceReport.EFERTDb
{
    [Table("AppUser")]
    public class AppUser
    {
        [Key]
        public string UserName { get; set; }
        
        public string Password { get; set; }

        public bool IsAdmin { get; set; }
    }
}
