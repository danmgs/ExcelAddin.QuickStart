using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using System.Linq;
using System.Windows.Forms;
using ExcelCore = Microsoft.Office.Core;

namespace ExcelAddin.QuickStart
{
    public partial class CustomRibbon
    {
        private void CustomRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void checkBoxShowTaskPanel_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CustomTaskPane.Visible = (sender as RibbonCheckBox).Checked;
        }

        private void switchPanelTab_Click(object sender, RibbonControlEventArgs e)
        {
            CustomTaskPane customTaskPane = Globals.ThisAddIn.CustomTaskPane;
            customTaskPane.Visible = true;
            TabControl tabMainControl = (TabControl)customTaskPane.Control.Controls.Find("tabMainControl", true).FirstOrDefault();

            if (tabMainControl.SelectedIndex == 0)
                tabMainControl.SelectTab(1);
            else
                tabMainControl.SelectTab(0);
        }

        private void toggleButtonPositionPanel_Click(object sender, RibbonControlEventArgs e)
        {
            CustomTaskPane customTaskPane = Globals.ThisAddIn.CustomTaskPane;
            if (customTaskPane.DockPosition == ExcelCore.MsoCTPDockPosition.msoCTPDockPositionRight)
                customTaskPane.DockPosition = ExcelCore.MsoCTPDockPosition.msoCTPDockPositionLeft;
            else
                customTaskPane.DockPosition = ExcelCore.MsoCTPDockPosition.msoCTPDockPositionRight;
        }
    }
}
