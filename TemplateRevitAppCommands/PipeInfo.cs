using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revit_Spec_VK
{
   public class PipeInfo
    {
       public string Type { get; set; }
       public string Size { get; set; }
       public string Mep_T_Wall { get; set; }

       public PipeInfo(string type, string size, string mep_t_wall)
       {
           this.Type = type;
           this.Size = size;
           this.Mep_T_Wall = mep_t_wall;

       }
    }
}
