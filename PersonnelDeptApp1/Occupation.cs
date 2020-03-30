using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PersonnelDeptApp1
{
    class Occupation
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Tarif { get; set; }

        public Occupation() { }
        public Occupation(int id, string name, decimal tarif = 0) {
            Id = id;
            Name = name;
            Tarif = tarif;
        }
    }
}
