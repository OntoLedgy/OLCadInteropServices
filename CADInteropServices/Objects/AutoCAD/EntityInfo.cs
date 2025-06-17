using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD
{
    public class EntityInfo
    {
        public string EntityType { get; set; }
        public string Handle { get; set; }
        public string Layer { get; set; }
        public string Color { get; set; }
        public string Linetype { get; set; }
        public double Lineweight { get; set; }
        public Dictionary<string, object> SpecificProperties { get; set; }

        public EntityInfo()
        {
            SpecificProperties = new Dictionary<string, object>();
        }
    }
}
