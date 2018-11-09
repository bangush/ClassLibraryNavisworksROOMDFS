using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Navisworks.Api.Plugins;

namespace ClassLibraryNavisworksROOMDFS
{
    [Plugin("ClassNavisRoomDFS",          //Plugin name
    "BIAD",                                 //4 character Developer ID or GUID
    ToolTip = "BIAD BIM STUDIO路径周游",                           //The tooltip for the item in the ribbon
    DisplayName = "BIAD BIM STUDIO路径周游")]                      //Display name for the Plugin in the Ribbon

    [DockPanePlugin(600, 500, AutoScroll = false, FixedSize = true)] // Default size

    public class ClassNavisRoomDFS : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            DFSControl control = null;

            // NOTE: All methods called from Navisworks should catch & handle 
            //       their own excepetions.
            try
            {
                control = new DFSControl();
                control.Dock = DockStyle.Fill;
                control.CreateControl();
                //control.RefreshTree();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name);
            }
            return control;
        }

        public override void DestroyControlPane(Control pane)
        {
            // NOTE: All methods called from Navisworks should catch handle 
            //       their own excepetions.
            try
            {
                DFSControl control = pane as DFSControl;
                if (control != null)
                {
                    control.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name);
            }
        }
    }
}
