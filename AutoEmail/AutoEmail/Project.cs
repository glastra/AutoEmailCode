using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoEmail
{
   public class Project
    {
        public Guid ProjectId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int ProjectType { get; set; }
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public long Cost { get; set; }
        public string AssignedTo { get; set; }


        public Project()
        {
            ProjectId = new Guid();
        }

        public Project(string name)
        {
            ProjectId = new Guid();
            Name = name;
        }




    }
}
