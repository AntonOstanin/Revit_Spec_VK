using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace Revit_Spec_VK
{
    public class ErrorMessage : IEquatable<ErrorMessage>
    {
     public string Message { get; set; }
     public int ID { get; set; }

      public ErrorMessage(string mes, ElementId id)
      {
          this.Message = mes;
          this.ID = id.IntegerValue;
      }

      public bool Equals(ErrorMessage other)
      {
          return ID==other.ID;
      }

      public int GetHashCode(ErrorMessage other)
      {
          return other.ID;
      }
    }
}
