using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace MyExcelAddIn.Controls
{
    public partial class CustomTaskPanelCtrl : UserControl
    {
        public CustomTaskPanelCtrl()
        {
            InitializeComponent();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void buttonLaunchAction_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Calling buttonLaunchAction_Click !");

            var excel = Globals.ThisAddIn.Application;
            var activeSheet = (Worksheet)excel.ActiveSheet;

            MessageBox.Show(string.Format("activeSheet.Name = {0}", activeSheet.Name));

            var cell = activeSheet.Range["B1", Type.Missing];
            cell.Value = "Hey buttonLaunchAction_Click !";

            //var cell2 = activeSheet.Range["A1", Type.Missing];
            //var mysheet = excel.ThisWorkbook.Sheets.Add(After: excel.ThisWorkbook.Sheets.Count);
        }

        public TabPage GetTabPage1()
        {
            return (TabPage) this.Controls.Find("tabPage1", true).First();
        }

        public TabPage GetTabPage2()
        {
            return (TabPage) this.Controls.Find("tabPage2", true).First();
        }
    }
}
