using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PersonnelDeptApp1
{
    class EmptyTableError : Exception
    {
        public EmptyTableError(string message) : base(message) {
        }
    }
}
