﻿using ExcelAddin.QuickStart.Controls;
using Microsoft.Office.Tools;
using System;
using System.Windows.Forms;

namespace ExcelAddin.QuickStart
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