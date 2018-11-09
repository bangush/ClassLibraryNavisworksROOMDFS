using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;


using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace ClassLibraryNavisworksROOMDFS
{
    [PluginAttribute("BIADBIMDFS",                 //Plugin name
    "BIAD",                                   //4 character Developer ID or GUID
    ToolTip = "BIAD BIM STUDIO路径周游",                 //The tooltip for the item in the ribbon
    DisplayName = "BIAD BIM STUDIO路径周游")]            //Display name for the Plugin in the Ribbon

    [AddInPluginAttribute(AddInLocation.AddIn,
                  LoadForCanExecute = true)]

    class ClassNavisDFS : AddInPlugin
    {
        public override CommandState CanExecute()
        {
            // NOTE: All methods called from Navisworks should catch handle 
            //       their own excepetions.
            

            try
            {
                //return new CommandState(false);
            }
            catch
            {
                return new CommandState(false);
            }
            return new CommandState(true);
        }

        public override int Execute(params string[] parameters)
        {
            //MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow, "Chris Chen");
            /*
            FormNavisworks f = new FormNavisworks();

            f.oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            f.Show();
           */
            try
            {
                //Find the plugin
                PluginRecord pr =
                   Autodesk.Navisworks.Api.Application.Plugins.FindPlugin("ClassNavisRoomDFS.BIAD");


                if (pr != null && pr is DockPanePluginRecord && pr.IsEnabled)
                {

                    //check if it needs loading
                    if (pr.LoadedPlugin == null)
                    {
                        pr.LoadPlugin();
                    }

                    DockPanePlugin dpp = pr.LoadedPlugin as DockPanePlugin;
                    if (dpp != null)
                    {
                        //switch the Visible flag
                        dpp.Visible = !dpp.Visible;
                    }
                }
                else
                {

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name);
            }

            return 0;
        }
    }
}
