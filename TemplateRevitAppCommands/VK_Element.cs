using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revit_Spec_VK
{
    public class VK_Element
    {
        public string CategoryName { get; private set; }
        public string Type { get; private set; }
        public string DetailType { get; private set; }
        public string BS_Name { get; private set; }
        public string Note { get; private set; }
        public string FamilyName { get; private set; }
        public string SystemType { get; private set; }

        public string FormulaParameter { get;  set; }

        public VK_Element(string catName, string type, string detType, string bs_name,string familyName, string formulaParameter, string note,string sysType)
        {
            this.CategoryName = catName;
            this.Type = type;
            this.DetailType = detType;
            this.BS_Name = bs_name;
            this.FormulaParameter = formulaParameter;
            this.Note = note;
            this.FamilyName = familyName;
            this.SystemType = sysType;
        }
    }




}
