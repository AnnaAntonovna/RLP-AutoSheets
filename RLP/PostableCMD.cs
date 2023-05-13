using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RLP
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LoadedTagsAndSymbols : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                PostCommandFast.ExecuteRevitCommand(commandData, PostableCommand.LoadedTagsAndSymbols);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Oh-oh, can't postCMD...");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
                return Result.Failed;
            }
        }
    }
    public static class PostCommandFast
    {
        public static void ExecuteRevitCommand(ExternalCommandData commandData, PostableCommand command)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(doc);
            UIApplication uiapp = new UIApplication(doc.Application);
            uiapp.PostCommand(RevitCommandId.LookupPostableCommandId(command));
        }
    }
}
