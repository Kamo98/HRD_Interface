using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PersonnelDeptApp1
{
    class Employee
    {
        public int Id{ get; set; }
        public string FIO { get; set; }

        public Employee() { }
        public Employee(int id, string fio) {
            Id = id;
            FIO = fio;
        }
    }
}
