using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wmi
{
    /// <summary>
    /// 
    /// </summary>
    internal class Binding
    {
        public string Name { get; set; }
        public string Filter { get; set; }
        public string Query { get; set; }
        public string Type { get; set; }
        public string Arguments { get; set; }
        public string Other { get; set; }

        public Binding(string name, string filter)
        {
            Name = name;
            Filter = filter;
        }
    }
}
