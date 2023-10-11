using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;
using System.IO;
using Environment = System.Environment;

namespace Vectra_SOW

{
    public partial class Form1 : Form
    {

        // Creating a Instance of all SOW Class
        Test1 test1 = new Test1();
        Test2 test2 = new Test2();
        Test3 test3 = new Test3();
        Test4 test4 = new Test4();
        Test5 test5 = new Test5();

        //Creating Instance
        CommonFunctions cmFunc = new CommonFunctions();




        public Form1()
        {
            //Intalizing Form
            InitializeComponent();
            
        }

        private void btTest1_Click(object sender, EventArgs e)
        {
            //Calling Test1 Main Function
             test1.Main();
            
        }

        private void btTest2_Click(object sender, EventArgs e)
        {
            string TextFile = null;

            //Getting Text File Path
            TextFile = cmFunc.GetTextFilePath();

            //Ensure Text file path is not Empty
            if (TextFile != null)
            {
                //Calling Test2 Main Function
                test2.Main(TextFile);
            }
            else
            {
                MessageBox.Show("Select a Text File & hit Run Button again to continue");
            }


        }

        private void btTest3_Click(object sender, EventArgs e)
        {
           
            test3.Main();
        }

        private void btTest4_Click(object sender, EventArgs e)
        {
            test4.Main();

        }

        private void btTest5_Click(object sender, EventArgs e)
        {

            string TextFile = null;

            //Getting Text File Path
            TextFile = cmFunc.GetTextFilePath();

            //Ensure Text file path is not Empty
            if (TextFile != null)
            {
                //Calling Test2 Main Function
                test5.Main(TextFile);
            }
            else
            {
                MessageBox.Show("Select a Text File & hit Run Button again to continue");
            }


        }

        private void btExit_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }
    }


}