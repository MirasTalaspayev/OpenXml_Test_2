using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml_Test_2
{
    class Document
    {
        public string Name { get; set; }
        public string Company { get; set; }
        public List<Position> positions = new List<Position>();
    }
}
