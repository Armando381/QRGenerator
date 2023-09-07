using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QRGenerator.Entities
{
    public class Student
    {
       public int Id { get; set; }
        public string? Name { get; set; }
        
        public int Age { get; set; }
        public int Grade { get; set; }
        public string? Gender { get; set; }
        public bool HasQr { get; set; } =false;
    }
}
