using ExcelAddin.QuickStart.Controls;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System;
using System.Deployment.Application;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

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

            //BindRandomValues(2, 3);
            AssignArrayToRange();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void BindRandomValues(int nbrows, int nbcols)
        {
            //((Worksheet)this.Application.ActiveWorkbook.Sheets[1]).Select();
            //Range r = this.Application.Worksheets[1] Range["B2", "B4"];


            //int startRow, endRow, startCol, endCol, row, col;
            //var singleData = new object[col];
            //var data = new object[row, col];
            ////For populating only a single row with 'n' no. of columns.
            //var startCell = (Range)worksheet.Cells[startRow, startCol];
            //startCell.Value2 = singleData;
            ////For 2d data, with 'n' no. of rows and columns.
            //var endCell = (Range)worksheet.Cells[endRow, endCol];
            //var writeRange = worksheet.Range[startCell, endCell];
            //writeRange.Value2 = data;

            object[,] workingValues = new object[nbrows, nbcols];

            for (int i = 0; i < nbrows; i++)
                for (int j = 0; j < nbcols; j++)
                    workingValues[i, j] = (i + 1) * (j + 1);

            Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)this.Application.Worksheets.Add();
            //Excel.Worksheet activeWorksheet = (Excel.Worksheet)Application.ActiveSheet;
            //Excel.Range r = activeWorksheet.get_Range("A1:F6", missing);

            Excel.Range r = newWorksheet.get_Range("A1:F6", missing);
            r.Value2 = workingValues;
        }

        /// <summary>
        /// Fast method for VSTO: assign to one defined range.
        /// https://stackoverflow.com/questions/3840270/fastest-way-to-interface-between-live-unsaved-excel-data-and-c-sharp-objects/3847094#3847094
        /// </summary>
        void AssignArrayToRange()
        {
            // Create the array.
            object[,] myArray = new object[10, 10];

            // Initialize the array.
            for (int i = 0; i < myArray.GetLength(0); i++)
            {
                for (int j = 0; j < myArray.GetLength(1); j++)
                {
                    myArray[i, j] = i + j;
                }
            }

            // Create a Range of the correct size:
            int rows = myArray.GetLength(0);
            int columns = myArray.GetLength(1);
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)Application.ActiveSheet;
            Excel.Range range = activeWorksheet.get_Range("A1", Type.Missing);
            range = range.get_Resize(rows, columns);

            // Assign the Array to the Range in one shot:
            range.set_Value(Type.Missing, myArray);
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
