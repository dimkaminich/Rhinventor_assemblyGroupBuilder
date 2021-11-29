using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rhinventor2021AssemblyGroupBuilder
{
    abstract class BaseComponent
    {
        public string compLabel { get; set; }
        public Rhino.Geometry.Plane compPlane { get; set; }
        public List<Dictionary<string, object>> compiPropertyData { get; set; }
    }
}
