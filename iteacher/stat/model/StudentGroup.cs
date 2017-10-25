using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iteacher.stat.model
{
    public class StudentGroup
    {
        
        public int Id { get; set; }
        public List<Student> Students { get; set; }

        public double Total { get; set; }

        public double Avg { get; set; }
    }
}
