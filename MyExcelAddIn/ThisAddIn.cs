using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using MyExcelAddIn.Controls;

namespace MyExcelAddIn
{
    public partial class ThisAddIn
    {
        //private UserControl1 myInputControl { get; set; }
        private Microsoft.Office.Tools.CustomTaskPane _customTaskPane { get; set; }
        public CustomRibbon _ribbon { get; set; }
        private CustomTaskPanelCtrl _customTaskPanelCtrl { get; set; }

        public CustomTaskPane CustomTaskPane
        {
            get
            {
                return _customTaskPane;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _customTaskPanelCtrl = new CustomTaskPanelCtrl();

            try
            {
                _customTaskPane = this.CustomTaskPanes.Add(_customTaskPanelCtrl, "#### My Task Panel ####");
                _customTaskPane.Width = 400;
                _customTaskPane.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
