using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITrainingPipeDiameterParam
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string outDiamParam = null;
            string inDiamParam = null;
            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));

            using(Transaction ts = new Transaction(doc, "Add parameter"))
            {
                ts.Start();
                CreateSharedParameter(uiapp.Application, doc, "Наименование", categorySet, BuiltInParameterGroup.PG_TEXT, true);
                ts.Commit();
            }

            List<Element> pipesList = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .ToList();

            foreach (Element element in pipesList)
            {
                Parameter outerDiameter = element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                outDiamParam = outerDiameter.AsValueString();
                Parameter innerDiameter = element.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                inDiamParam = innerDiameter.AsValueString();

                using(Transaction ts2 = new Transaction(doc, "SetParameters"))
                {
                    ts2.Start();
                    var familyInstance = element as FamilyInstance;
                    Parameter name = familyInstance.LookupParameter("Наименование");
                    name.Set($"Труба {outDiamParam}/{inDiamParam}");
                    ts2.Commit();
                }
            }
            
            TaskDialog.Show("Сообщение", "Параметр добавлен");
            return Result.Succeeded;
        }

        private void CreateSharedParameter(Application application, Document doc, 
            string parameterName, CategorySet categorySet, BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile defFile = application.OpenSharedParameterFile();
            if (defFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = defFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        }
    }
}
