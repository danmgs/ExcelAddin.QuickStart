using ExcelAddin.QuickStart.Controls;
using Microsoft.Office.Tools;
using System;
using System.Deployment.Application;
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
            //https://stackoverflow.com/questions/24282560/get-office-addin-publish-version
            string versionNumber = string.Empty;
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment applicationDeployment = ApplicationDeployment.CurrentDeployment;
                Version version = applicationDeployment.CurrentVersion;
                versionNumber = String.Format("{0}.{1}.{2}.{3}", version.Major, version.Minor, version.Build, version.Revision);
            }

            _customTaskPanelCtrl = new CustomTaskPanelCtrl(versionNumber);

            try
            {
                _customTaskPane = this.CustomTaskPanes.Add(_customTaskPanelCtrl, "#### My Task Panel ####");
                _customTaskPane.Width = 400;
                _customTaskPane.Visible = true;


                //labelVersionTab1.Text = string.Format("Addin v{0}", versionNumber);
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
