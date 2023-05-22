using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace RLP
{
    public class TabPanel : IExternalApplication
    {
        string assemPath;
        string path;
        RibbonPanel panel1;
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // Names of the tab, panel and buttons!

            string tabName = "RLP - Сборный железобетон";
            string panelName1 = "Cборный железобетон";

            //Creating the tab 

            application.CreateRibbonTab(tabName);

            //Creating panels

            panel1 = application.CreateRibbonPanel(tabName, panelName1);

            //Things that will help us later

            Assembly assembly = Assembly.GetExecutingAssembly();
            assemPath = assembly.Location;
            string assebmlyDir = new FileInfo(assembly.Location).DirectoryName;
            path = System.IO.Path.GetDirectoryName(assemPath);

            //Create buttonsCreateButton

            CreateButton("Виды и листы", 1, $"RLP.Combine", panel1, "Выберите сборку и кнопка создаст листы и фасады к ней");
            CreateButton("Проставление\nмарок", 2, $"RLP.Annotations", panel1, "Проставит марки для элементов сборки на активном виде (в прогрессе)");
            CreateButton("Изменить марки\nпо умолчанию", 3, $"RLP.LoadedTagsAndSymbols", panel1, "Загруженные в проект марки");
            CreateButton("Создать\nспецификации", 4, $"RLP.ScheduleCMD", panel1, "Загруженные в проект марки");
            //CreateButton("Тест", 3, $"RLP.CreateAssemblySection", panel1, "Wow - a tooltip!");

            return Result.Succeeded;
        }

        private void CreateButton(string buttonName, int index, string cmdName)
        {
            string inButtinName = $"Button {index}";
            //comands and tips should be added by list
            PushButtonData pushButtonData = new PushButtonData(inButtinName, buttonName, assemPath, cmdName);
            pushButtonData.LargeImage = new BitmapImage(new Uri(Path.Combine(path, $@"Images\icon_{index}.png")));
            panel1.AddItem(pushButtonData);
        }

        private void CreateButton(string buttonName, int index, string cmdName, RibbonPanel panel)
        {
            string inButtinName = $"Button {index}";
            //comands and tips should be added by list
            PushButtonData pushButtonData = new PushButtonData(inButtinName, buttonName, assemPath, cmdName);
            pushButtonData.LargeImage = new BitmapImage(new Uri(Path.Combine(path, $@"Images\icon_{index}.png")));
            panel.AddItem(pushButtonData);
        }
        private void CreateButton(string buttonName, int index, string cmdName, RibbonPanel panel, string toolTip)
        {
            string inButtinName = $"Button {index}";
            //comands and tips should be added by list
            PushButtonData pushButtonData = new PushButtonData(inButtinName, buttonName, assemPath, cmdName);
            pushButtonData.LargeImage = new BitmapImage(new Uri(Path.Combine(path, $@"Images\icon_{index}.png")));
            pushButtonData.ToolTip = toolTip;
            panel.AddItem(pushButtonData);
        }
    }
}
