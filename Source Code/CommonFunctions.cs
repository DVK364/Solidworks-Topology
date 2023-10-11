
using SolidWorks.Interop.sldworks;
using Environment = System.Environment;

namespace Vectra_SOW
{
    public class CommonFunctions
    {
        // Function to Show message Box with Cancel Button
        public void MsgBoxWithCancelBt(string msgContent, string msgHeader)
        {

           DialogResult dialogResult = MessageBox.Show(msgContent, msgHeader, MessageBoxButtons.OKCancel);

            if (dialogResult == DialogResult.Cancel)
            {
                Environment.Exit(0);
            }
        }

        //Function to Clear all existing selection filters (selectable Entites like Edges, Face, Point & So on)
        public void ClearSelFilters(SldWorks swApp)
        {

            // Get all exixting Selection Filters
            object paramsObject = swApp.GetSelectionFilters();

            // Clear all exixting Selection Filters
            swApp.SetSelectionFilters(paramsObject, false);


        }

        //Function to get Text file path
        public string GetTextFilePath()
        {

            string TextFile = null;

            //Calling file dialog to get text file
            OpenFileDialog bDialog = new OpenFileDialog();

            //Give Title for selection
            bDialog.Title = "Select Text File";

            //Set Selection Filter to select only Text File
            bDialog.Filter = "Text Files|*.txt";

            //Set Inital File open Directory as C Drive
            bDialog.InitialDirectory = @"C:\";

            //Ensure file dialog has been Intiated 
            if (bDialog.ShowDialog() == DialogResult.OK)
            {
                //Get selected file path to Proceed.
                TextFile = bDialog.FileName;
            }

            return TextFile;

        }




    }
}
