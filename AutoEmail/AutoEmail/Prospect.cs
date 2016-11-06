using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoEmail
{
    public class Prospect
    {
        public Guid ProspectId { get; set; }
        public String Name { get; set; }
        public String LastName { get; set; }
        public string Email { get; set; }
        public DateTime ProspectDateTime { get; set; }
        public DateTime LastContactDateTime { get; set; }
    }
}
