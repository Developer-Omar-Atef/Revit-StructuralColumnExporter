using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace KaitechColumnsReportAddin
{
    public class App : IExternalApplication
    {
        static string AddInPath = typeof(App).Assembly.Location;

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }


        public Result OnStartup(UIControlledApplication application)
        {

            string tabName = "KAITECH-BD-R06";
            string panelName = "Structure";
            string buttonName = "ColumnsReport"; // Internal name for PushButtonData
            string buttonText = "Columns Report"; // Displayed text on the button

            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch (Exception) { /* Tab already exists */ }

            RibbonPanel panel = null;
            var panels = application.GetRibbonPanels(tabName);
            foreach (var p in panels)
            {
                if (p.Name == panelName)
                {
                    panel = p;
                    break;
                }
            }
            if (panel == null)
            {
                panel = application.CreateRibbonPanel(tabName, panelName);
            }

            PushButtonData pbd = new PushButtonData(
                buttonName,
                buttonText,
                AddInPath, // Assembly path
                typeof(Command).FullName // Full class name for the command
            );

            // The resource names must match: YourProjectName.FolderName.FileName.Extension
            string assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
            try
            {
                pbd.LargeImage = PngImageSource(assemblyName + ".Resources.report_icon_32.png");
                pbd.Image = PngImageSource(assemblyName + ".Resources.report_icon_16.png");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Icon Error", "Could not load ribbon icons: " + ex.Message);
            }

            pbd.ToolTip = "Exports a report of structural columns from the Revit model.";

            panel.AddItem(pbd);

            return Result.Succeeded;
        }

        private System.Windows.Media.ImageSource PngImageSource(string embeddedPath)
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(embeddedPath);
            PngBitmapDecoder decoder = new PngBitmapDecoder(stream,BitmapCreateOptions.PreservePixelFormat,BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}
