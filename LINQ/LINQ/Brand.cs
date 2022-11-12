using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LINQ
{
    class Brand
    {
        public int ID { get; set; }
        public string Name { get; set; }

        public static implicit operator Brand(string v)
        {
            throw new NotImplementedException();
        }
    }
}
