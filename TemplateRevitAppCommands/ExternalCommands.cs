using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Plumbing;
using System.Reflection;
using System.Windows.Forms;
using Autodesk.Revit.DB.Mechanical;
using Revit_Spec_VK.PluginStatTableAdapters;
using System.Diagnostics;
using Autodesk.Revit.DB.Structure;

namespace Revit_Spec_VK
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class ExternalCommands : IExternalCommand
    {
        public static List<ErrorMessage> errorMessages = new List<ErrorMessage>();

        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            if (MainForm.mf != null)
                MainForm.mf.Close();
            errorMessages = new List<ErrorMessage>();
            FrameWork fw = new FrameWork();
            UIApplication appRevit = commandData.Application;
            Document doc = appRevit.ActiveUIDocument.Document;
            if (Environment.UserName != "ostaninam")
            {
                try
                {
                    C_PluginStatisticTableAdapter pg = new C_PluginStatisticTableAdapter();
                    pg.Insert("Revit", "Спецификация ВК", FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion,
                        appRevit.ActiveUIDocument.Document.Title,
                        Environment.UserName, DateTime.Now);
                }
                catch { }
            }
            string pathToSpecExcel = @"\\dsk2.picompany.ru\project\CAD_Settings\Revit_server\03. Project Templates\04. Водоснабжение_канализация\Спецификация\Спецификации_ВК.xlsx";
            string pathToSpecRDExcel = @"\\dsk2.picompany.ru\project\CAD_Settings\Revit_server\03. Project Templates\04. Водоснабжение_канализация\Спецификация\Спецификации_R&D.xlsx";
            string pathToUserExcel = @"\\dsk2.picompany.ru\project\CAD_Settings\Revit_server\03. Project Templates\04. Водоснабжение_канализация\Спецификация\Users_Spec_VK.xlsx";
            string localPathExcel = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Спецификации_ВК.xlsx";
            string localPathUserExcel = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\ЮЗвери_ВК.xlsx";
            //    File.Copy(pathToSpecExcel, localPathExcel, true);
            File.Copy(pathToUserExcel, localPathUserExcel, true);
            var usersRD = fw.GetUsersRD(localPathUserExcel);
            if (usersRD.Any(x => x.Equals(Environment.UserName.ToLower())))
                pathToSpecExcel = pathToSpecRDExcel;
            File.Copy(pathToSpecExcel, localPathExcel, true);
            var vkExcelElements = fw.GetVkElements(localPathExcel).GroupBy(x => x.CategoryName).Select(x => new { CategoryName = x.Key, VkElements = x.ToList() }).ToList();//.GroupBy(x => x.CategoryName).Select(g => g.ToList()).ToList();
            var allElements = new FilteredElementCollector(doc).WhereElementIsNotElementType().Where(x => (x.Category != null)).ToList();
            var neededElements = allElements.Where(element => vkExcelElements.Any(x => x.CategoryName.Equals(element.Category.Name))).GroupBy(x => x.Category.Name).Select(x => new
             {
                 CategoryName = x.Key,
                 Elements = x.ToList()
             }).ToList();

            var paramCatInfos = fw.GetParameterCategories(localPathExcel);





            var allPypes = allElements.Where(x => x.Category.Name.Equals("Трубы")).ToList();

            var pipeInfos = fw.GetPipeInfo(localPathExcel);








            using (Transaction tr = new Transaction(doc, "Спецификация ВК"))
            {
                tr.Start();

                #region Заполнение Параметров по значениям
                SetParametersByValue(doc, allElements, paramCatInfos);
                #endregion

                #region Заполнение Mep_t_стенки
                SetMep_t(allPypes, pipeInfos);
                #endregion

                #region Заполнение PIC_Поставщик
                SetPIC_Creator(doc, allElements);
                #endregion




                foreach (var categoryNeededEl in neededElements)            //Категории элементов в модели
                {
                    foreach (var needEl in categoryNeededEl.Elements)       //Перебор элементов категории
                    {
                        var excelCategory =
                            vkExcelElements.First(x => x.CategoryName.Equals(categoryNeededEl.CategoryName));      //Необходимая категория из Excel
                        List<VK_Element> vkOverlap = new List<VK_Element>();
                        foreach (var vkEl in excelCategory.VkElements)
                        {
                            if (vkEl.Type != "" && !vkEl.Type.Equals(needEl.Name))        //Тип
                            {
                                if (vkEl.CategoryName == "Трубы" &
                                    !excelCategory.VkElements.Any(x => x.Type.Equals(needEl.Name)))
                                {
                                    errorMessages.Add(
                                        new ErrorMessage(string.Format("Тип трубы не найден в файле Excel."), needEl.Id));
                                    break;
                                }

                                continue;
                            }


                            FamilyInstance fi = needEl as FamilyInstance;                        //Тип детали
                            if (fi != null)
                            {
                                MechanicalFitting mechFit = fi.MEPModel as MechanicalFitting;
                                if (mechFit != null && vkEl.DetailType != "" && !vkEl.DetailType.Equals(mechFit.PartType.ToString()))
                                    continue;
                            }

                            FamilySymbol fs = doc.GetElement(needEl.GetTypeId()) as FamilySymbol;

                            if (fs != null && Util.SkipCategories.Any(x => x.CategoryName.Equals(categoryNeededEl.CategoryName) & x.FamilyName.Equals(fs.FamilyName)))
                                break;                        //Пропуск элементов из списка "Не обрабатывать"!!!!!!!


                            if (vkEl.BS_Name != "" && !vkEl.BS_Name.Equals(GetParameter(needEl, "BS_Наименование")))     //BS_Наименование
                                continue;


                            if (vkEl.FamilyName != "")           //Имя семейства
                            {
                                if (fs != null)
                                {
                                    if (!vkEl.FamilyName.Equals(fs.FamilyName))
                                        continue;
                                }
                            }
                            vkOverlap.Add(vkEl);
                        }


                        if (vkOverlap.Count > 1)
                        {
                            vkOverlap = vkOverlap.Where(x => x.BS_Name != "" | x.FamilyName != "" | x.FormulaParameter.Equals(" ")).ToList();
                            if (vkOverlap.Count > 1)
                                vkOverlap = vkOverlap.Where(x => x.FamilyName != "").ToList();
                        }


                        if (vkOverlap.Count == 0)
                            continue;

                        GetParameterByFormula(vkOverlap, needEl);

                    }
                }
                tr.Commit();
            }
            if (errorMessages.Count != 0)
            {
                MainForm mf = new MainForm(appRevit, errorMessages.OrderBy(x => x.Message).ToList());
                mf.Show();
            }
            else
            {
                MessageBox.Show("Выполнено без ошибок!", "Спецификация ИОС", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            return Result.Succeeded;
        }

        private static void SetMep_t(List<Element> allPypes, List<PipeInfo> pipeInfos)
        {
            foreach (var pipe in allPypes)
            {
                string type = pipe.Name;
                string size = pipe.LookupParameter("Размер").AsString();
                var p = pipeInfos.Where(x => x.Type.Equals(type) && x.Size.Equals(size)).ToList();
                if (p.Count == 0)
                    continue;

                if (pipe.LookupParameter("MEP_t_стенки") == null) continue;
                pipe.LookupParameter("MEP_t_стенки").Set(p[0].Mep_T_Wall);
            }
        }

        private void SetPIC_Creator(Document doc, List<Element> allElements)
        {
            foreach (var el in allElements)
            {
                FamilyInstance fi = el as FamilyInstance;
                if (fi == null) continue;
                Parameter par = doc.GetElement(fi.GetTypeId()).LookupParameter("PIC_Поставщик");
                Parameter parCreator = doc.GetElement(fi.GetTypeId()).LookupParameter("Изготовитель");
                if (par == null || parCreator == null)
                    continue;
                string value = GetParameterTypeValue(par, false);
                if (value != "") continue;
                par.Set(GetParameterTypeValue(parCreator, false));

            }
        }

        private void SetParametersByValue(Document doc, List<Element> allElements, List<ParameterCategory> paramCatInfos)
        {
            FrameWork fw = new FrameWork();
            var neededCatParamInfos = allElements.Where(element => paramCatInfos.Any(x => x.CategoryName.Equals(element.Category.Name))).GroupBy(x => x.Category.Name).Select(x => new
            {
                CategoryName = x.Key,
                Elements = x.ToList()
            }).ToList();

            foreach (var categoryParam in neededCatParamInfos)
            {
                foreach (var elementCategory in categoryParam.Elements)
                {
                    string cat = categoryParam.CategoryName;
                    string type = doc.GetElement(elementCategory.GetTypeId()).Name;
                    string famName = "";
                    FamilySymbol fs = null;
                    if (elementCategory is FamilyInstance)
                    {
                        fs = (elementCategory as FamilyInstance).Symbol;
                        famName = fs.FamilyName;
                    }
                    bool isSetParameter = false;
                    foreach (var filterParam in paramCatInfos.Where(x => x.CategoryName.Equals(cat)).ToList())
                    {
                        if (!filterParam.TypeName.Equals("") & !type.Equals(filterParam.TypeName))
                            if (filterParam.TypeName.Contains("*") & !(type.Contains(filterParam.TypeName.Replace("*", ""))))
                                continue;
                            else if (!filterParam.TypeName.Contains("*"))
                                continue;

                        if (!filterParam.FamilyName.Equals("") & !famName.Equals(filterParam.FamilyName))
                            if (filterParam.FamilyName.Contains("*") & !(famName.Contains(filterParam.FamilyName.Replace("*", ""))))
                                continue;
                            else if (!filterParam.FamilyName.Contains("*"))
                                continue;
                        bool isValid = true;
                        foreach (var p in filterParam.parameters)
                        {
                            if (p.Value.Equals("")) continue;
                            Parameter par = elementCategory.LookupParameter(p.Name);
                            if (par == null)
                                par = doc.GetElement(elementCategory.GetTypeId()).LookupParameter(p.Name);
                            if (par == null)
                                continue;
                            string parValue = fw.GetParameterValue(par);
                            if (p.Value.Contains("*") & parValue.Contains(p.Value.Replace("*", "")))
                                continue;
                            if (parValue == p.Value)
                                continue;
                            isValid = false;
                            break;

                        }
                        if (!isValid)
                        {
                            //Не произошло совпадение
                            continue;
                        }
                        Parameter paramToSet = elementCategory.LookupParameter(filterParam.ParameterToSet);
                        if (paramToSet == null)
                        {
                            if (fs != null)
                                paramToSet = fs.LookupParameter(filterParam.ParameterToSet);
                            if (paramToSet == null)
                            {
                                errorMessages.Add(new ErrorMessage(string.Format("Лист Параметры. Отсутствует параметр \"{0}\"", filterParam.ParameterToSet), elementCategory.Id));
                                continue; //Нет заполняемого параметра
                            }
                        }
                        AnalyticalModelSurface sur = null;

                        isSetParameter = true;
                        if (paramToSet.IsReadOnly)
                        {
                            errorMessages.Add(new ErrorMessage(string.Format("Лист Параметры. Значение параметра \"{0}\" задано формулой в семействе. ", filterParam.ParameterToSet), elementCategory.Id));
                            continue;
                        }
                        paramToSet.Set(filterParam.ParameterValue);
                        break;
                    }
                    if (!isSetParameter)
                    {
                        //Элемент не найден в Excel
                    }
                }

            }
        }

        private void GetParameterByFormula(List<VK_Element> vkOverlap, Element element)
        {
            string[] masParams = vkOverlap.First().FormulaParameter.Split('|');
            string paramValue = "";
            bool isBreak = false;
            foreach (var str in masParams)
            {
                if (str.Contains("*"))
                {
                    string valParam = GetParameter(element, str.Remove(0, 1));
                    if (str.Equals("*Тип"))
                        valParam = element.Document.GetElement(element.GetTypeId()).Name;
                    if (element is PipeInsulation)                                                        //Изоляция
                    {
                        PipeInsulation pipeInsul = element as PipeInsulation;
                        Pipe pipe = element.Document.GetElement(pipeInsul.HostElementId) as Pipe;

                        if (pipe == null)
                        {
                            isBreak = true;
                            break;
                        }

                        if (str.Equals("*Внешний диаметр"))
                        {
                            valParam = GetParameter(pipe, str.Remove(0, 1));

                            if (vkOverlap.First().Note.Contains("INSULATION"))
                            {
                                string thicknes = Math.Round(pipeInsul.Thickness * 304.8, 0).ToString() + " мм";
                                var ins = Util.Insulations.Where(x => x.SizeInsulation.Equals(thicknes)).ToList();
                                if (ins.Count != 0)
                                {
                                    var diametr = ins[0].Diameter.Where(x => x.Equals(Convert.ToInt16(double.Parse(valParam)))).ToList();
                                    if (diametr.Count == 0)
                                    {

                                        for (int i = 0; i < ins[0].Diameter.Count - 2; i++)
                                        {
                                            if (!(ins[0].Diameter[i] < Convert.ToInt16(double.Parse(valParam)) &
                                                  ins[0].Diameter[i + 1] > Convert.ToInt16(double.Parse(valParam)))) continue;
                                            valParam = ins[0].Diameter[i + 1].ToString();
                                            break;
                                        }
                                        diametr = ins[0].Diameter.Where(x => x.Equals(Convert.ToInt16(double.Parse(valParam)))).ToList();
                                        if (diametr.Count == 0) valParam = ins[0].Diameter[0].ToString();

                                    }
                                    else valParam = diametr[0].ToString();
                                }
                                else
                                {
                                    errorMessages.Add(new ErrorMessage(string.Format("Значение толщины изоляции, которое отсутствует на листе INSULATION"), pipeInsul.Id));
                                    paramValue = "";
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {

                        if (str.Equals("*Размер"))
                        {

                            if (vkOverlap.First().Note.Contains("UNIQUE SIZE"))
                            {
                                string newVal = "";
                                string[] masVal = valParam.Split('x');
                                for (int i = 0; i < masVal.Length; i++)
                                {
                                    if (newVal.Contains(masVal[i]))
                                        continue;
                                    newVal += masVal[i] + "x";
                                }
                                if (valParam.Length > 1)
                                    valParam = newVal.Remove(newVal.Length - 1, 1);
                            }
                            if (vkOverlap.First().Note.Contains("THREAD"))
                            {
                                string newVal = "";
                                string[] masVal = valParam.Split('x');
                                int minVal = masVal.Min(x => Convert.ToInt32(x));
                                for (int i = 0; i < masVal.Length; i++)
                                {

                                    var f = Util.VkThreads.Where(x => x.Millimeter.Equals(minVal.ToString())).ToList();
                                    //Соответствие миллиметры - дюймы
                                    if (f.Count == 0 | masVal[i] != minVal.ToString())
                                    {
                                        newVal += masVal[i] + "x";
                                        continue;
                                    }
                                    newVal += f[0].Inch + "x";
                                }
                                valParam = newVal.Remove(newVal.Length - 1, 1);
                            }

                        }
                    }
                    paramValue += valParam;
                }
                else paramValue += str;

            }
            if (!isBreak)
            {
                if (element.LookupParameter(VkParameter.PikNameByGost) != null)
                    element.LookupParameter(VkParameter.PikNameByGost).Set(paramValue);
                else
                {
                    FamilySymbol fs = element.Document.GetElement(element.GetTypeId()) as FamilySymbol;
                    string name = element.Document.GetElement(element.GetTypeId()).Name;
                    if (fs != null)
                    {
                        name = fs.FamilyName;
                    }
                    errorMessages.Add(new ErrorMessage("Отсутствует параметр «PIC_Наименование_по_ГОСТ» " + name, element.Id));
                }

            }
        }

        string GetParameter(Element element, string nameParameter)
        {
            string paramValue = "";
            var parameter = element.LookupParameter(nameParameter) ??
                            element.Document.GetElement(element.GetTypeId()).LookupParameter(nameParameter);
            FamilySymbol fs = element.Document.GetElement(element.GetTypeId()) as FamilySymbol;
            string nameType = element.Document.GetElement(element.GetTypeId()).Name;
            if (parameter == null)
            {


                if (fs != null)
                {
                    nameType = fs.FamilyName;
                }
                if ((nameParameter == "Внешний диаметр" && element is PipeInsulation))// | (nameParameter == "BS_Наименование" && nameType.ToUpper().StartsWith("MEQ_ШПК")))
                    return "";
                errorMessages.Add(new ErrorMessage(string.Format("Отсутствует параметр «{0}» " + nameType, nameParameter), element.Id));
                return "";
            }
            bool isCorner = nameParameter.ToUpper().Contains("УГОЛ");
            string val = GetParameterTypeValue(parameter, isCorner);

            if (val == "" & nameParameter != "Тип")
            {

                if (fs != null)
                {
                    nameType = fs.FamilyName;
                }
                errorMessages.Add(new ErrorMessage(string.Format("Пустое значение у параметра «{0}» " + nameType, nameParameter), element.Id));
            }
            return val ?? (val = "");
        }

        string GetParameterTypeValue(Parameter parameter, bool isCorner)
        {
            string value = "";
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    if (isCorner) value = Math.Round(parameter.AsDouble() * 180 / 3.14, 1).ToString();
                    else
                        value = Math.Round((parameter.AsDouble() * 304.8), 1).ToString(); break;
                case StorageType.Integer: value = Math.Round((parameter.AsInteger() * 304.8), 0).ToString(); break;
                case StorageType.String: value = Convert.ToString(parameter.AsString()); break;
            }
            if (value == null)
                return "";
            return value;
        }

    }
}
