using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using OfficeOpenXml.Drawing.Chart;

namespace Revit_Spec_VK
{
    public class FrameWork
    {

        public string GetParameterValue(Parameter par)
        {
            string value = par.AsValueString();
            switch (par.StorageType)
            {
                case StorageType.Double:
                    value = Math.Round(par.AsDouble()*304.8,1).ToString();                   
                    break;

                //case StorageType.ElementId:
                //    ElementId id = t.AsElementId(fp);
                //    Element e = doc.get_Element(ref id);
                //    value = id.Value.ToString() + " ("
                //      + Util.ElementDescription(e) + ")";
                //    break;

                case StorageType.Integer:
                    value = par.AsInteger().ToString();
                    break;

                case StorageType.String:
                    value = par.AsString();
                    break;
            }
            return value;
        }

        public List<string> GetUsersRD(string pathToExcel)
        {
            List<string> users = new List<string>();
            using (var xlPackage = new ExcelPackage(new FileInfo(pathToExcel)))
            {
                var workBook = xlPackage.Workbook;
                var ws = workBook.Worksheets[1];
                int row = 2;
                while (ws.Cells[row, 1].Value != null)
                {
                    users.Add(Convert.ToString(ws.Cells[row, 1].Value).Trim().ToLower());
                    row++;
                }
            }
            return users;
        }

        public List<ParameterCategory> GetParameterCategories(string pathToExcel)
        {
            List<ParameterCategory> paramCatInfos = new List<ParameterCategory>();
            using (var xlPackage = new ExcelPackage(new FileInfo(pathToExcel)))
            {
                var workBook = xlPackage.Workbook;
                var ws = workBook.Worksheets["Параметры"];
                int row = 2;
                while (ws.Cells[row, 1].Value != null)
                {
                    string paramToSet = Convert.ToString(ws.Cells[row, 1].Value).Trim();
                    string paramValue = Convert.ToString(ws.Cells[row, 2].Value).Trim();
                    string cat = Convert.ToString(ws.Cells[row, 3].Value).Trim();
                    string type = Convert.ToString(ws.Cells[row, 4].Value).Trim();
                    string famName = Convert.ToString(ws.Cells[row, 5].Value).Trim();
                    var paramCat = new ParameterCategory(cat, famName, type, paramToSet, paramValue);
                    int columnCount = 6;
                    while (Convert.ToString(ws.Cells[1, columnCount].Value).Trim() != "")
                    {
                        paramCat.parameters.Add(new VK_Parameter(Convert.ToString(ws.Cells[1, columnCount].Value).Trim(), Convert.ToString(ws.Cells[row, columnCount].Value).Trim()));

                        columnCount++;
                    }
                    paramCatInfos.Add(paramCat);
                    row++;
                }
            }
            return paramCatInfos;
        }

        public List<PipeInfo> GetPipeInfo(string pathToExcel)
        {
            List<PipeInfo> pipeInfos = new List<PipeInfo>();
            using (var xlPackage = new ExcelPackage(new FileInfo(pathToExcel)))
            {
                var workBook = xlPackage.Workbook;
                var ws = workBook.Worksheets["MEP_t_стенки"];
                int row = 2;
                while (ws.Cells[row, 1].Value != null)
                {
                    string type = Convert.ToString(ws.Cells[row, 1].Value).Trim();
                    string size = Convert.ToString(ws.Cells[row, 2].Value).Trim();
                    string mep_t = Convert.ToString(ws.Cells[row, 3].Value).Trim();
                    pipeInfos.Add(new PipeInfo(type, size, mep_t));
                    row++;
                }
            }
            return pipeInfos;
        }

        public List<VK_Element> GetVkElements(string pathToExcel)
        {
            List<VK_Element> vkElements = new List<VK_Element>();

            using (var xlPackage = new ExcelPackage(new FileInfo(pathToExcel)))
            {
                var workBook = xlPackage.Workbook;
                var ws = workBook.Worksheets[1];
                int row = 2;
                while (ws.Cells[row, 1].Value != null)
                {
                    // if (vkElements.Any(x => x.BS_Name.Equals()))
                    string cat = Convert.ToString(ws.Cells[row, 1].Value).Trim();
                    string type = Convert.ToString(ws.Cells[row, 2].Value).Trim();
                    string typeDet = Convert.ToString(ws.Cells[row, 3].Value).Trim();
                    string bsName = Convert.ToString(ws.Cells[row, 4].Value).Trim();
                    string fam = Convert.ToString(ws.Cells[row, 5].Value).Trim();
                    string form = Convert.ToString(ws.Cells[row, 6].Value).Trim();
                    string note = Convert.ToString(ws.Cells[row, 7].Value).Trim();

                    if (vkElements.Any(x =>
                        x.CategoryName.Equals(cat) & x.Type.Equals(type) & x.DetailType.Equals(typeDet) &
                        x.BS_Name.Equals(bsName) & x.FamilyName.Equals(fam)))
                    {
                        ExternalCommands.errorMessages.Add(new ErrorMessage(string.Format("В Excel дублируется строка:" +
                                                                                          "(Категория: \"{0}\",Тип: \"{1}\",Тип детали: \"{2}\",BS_: \"{3}\",Семейство: \"{4}\")",
                                                                                          cat, type, typeDet, bsName, fam), new ElementId(0)));
                        vkElements.First(x =>
                            x.CategoryName.Equals(cat) & x.Type.Equals(type) & x.DetailType.Equals(typeDet) &
                            x.BS_Name.Equals(bsName) & x.FamilyName.Equals(fam)).FormulaParameter = " ";
                        form = " ";
                    }
                    else vkElements.Add(new VK_Element(cat, type, typeDet, bsName, fam, form, note));
                    //  continue;


                    row++;
                }
                row = 2;
                Util.VkThreads = new List<VkThread>();
                ws = workBook.Worksheets["THREAD"];
                while (ws.Cells[row, 1].Value != null)
                {
                    Util.VkThreads.Add(new VkThread(Convert.ToString(ws.Cells[row, 1].Value).Trim(), Convert.ToString(ws.Cells[row, 2].Value).Trim()));
                    row++;
                }


                int column = 1;
                Util.Insulations = new List<Insulation>();
                ws = workBook.Worksheets["INSULATION"];
                while (ws.Cells[1, column].Value != null)
                {
                    string sizeInsul = Convert.ToString(ws.Cells[1, column].Value).Trim();
                    Util.Insulations.Add(new Insulation(sizeInsul));
                    row = 2;
                    while (ws.Cells[row, column].Value != null)
                    {
                        Util.Insulations.First(x => x.SizeInsulation.Equals(sizeInsul)).Diameter.Add(Convert.ToInt16(ws.Cells[row, column].Value));
                        row++;
                    }
                    column++;
                }

                Util.SkipCategories = new List<SkipCategory>();
                ws = workBook.Worksheets["Не обрабатывать"];
                row = 2;
                while (ws.Cells[row, 1].Value != null)
                {
                    Util.SkipCategories.Add(new SkipCategory(Convert.ToString(ws.Cells[row, 1].Value).Trim(), Convert.ToString(ws.Cells[row, 2].Value).Trim()));
                    row++;
                }


            }
            return vkElements;
        }
    }
}
