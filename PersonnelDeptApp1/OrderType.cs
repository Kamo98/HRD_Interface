using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PersonnelDeptApp1
{
    class OrderType
    {
        public int Id { get; private set; }
        public string Name { get; private set; }

        public OrderType(int id, string name) {
            Id = id;
            Name = name;
        }

    }
}
