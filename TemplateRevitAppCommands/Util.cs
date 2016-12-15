using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revit_Spec_VK
{
    public static class Util
    {
        public static List<VkThread> VkThreads { get; set; }
        public static List<Insulation> Insulations { get; set; }
        public static List<SkipCategory> SkipCategories { get; set; }
        public static List<ParameterCategory> ParameterCategories { get; set; }

    }

    public class VkThread
    {
        public string Millimeter { get; set; }
        public string Inch { get; set; }

        public VkThread(string mm, string inch)
        {
            this.Millimeter = mm;
            this.Inch = inch;
        }
    }


    public class ParameterCategory : ICategory
    {
        public string CategoryName { get; set; }

        public string ParameterToSet { get; set; }

        public string ParameterValue { get; set; }
        public string TypeName { get; set; }
        public string FamilyName { get; set; }
        public List<VK_Parameter> parameters { get; set; }


        public ParameterCategory(string catName, string famName, string typeName, string paramToSet, string paramValue)
        {
            this.CategoryName = catName;
            this.TypeName = typeName;
            this.FamilyName = famName;
            this.ParameterToSet = paramToSet;
            this.ParameterValue = paramValue;
            parameters = new List<VK_Parameter>();
        }
    }

    public class VK_Parameter
    {
        public string Name { get; set; }
        public string Value { get; set; }


        public VK_Parameter(string name, string value)
        {
            this.Name = name;
            this.Value = value;
        }
    }

    public class SkipCategory : ICategory
    {
        public string CategoryName { get; set; }
        public string FamilyName { get; set; }

        public SkipCategory(string catName, string famName)
        {
            this.CategoryName = catName;
            this.FamilyName = famName;
        }
    }

    public class Insulation
    {
        public string SizeInsulation { get; set; }
        public List<int> Diameter { get; set; }
        public Insulation(string sizeInsul)
        {
            this.SizeInsulation = sizeInsul;
            Diameter = new List<int>();
        }
    }
}
