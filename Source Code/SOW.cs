using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using View = SolidWorks.Interop.sldworks.View;
using System.Runtime.InteropServices;
using Attribute = SolidWorks.Interop.sldworks.Attribute;



namespace Vectra_SOW
{
    
    public class Test1
    {
        CommonFunctions cmFunc = new CommonFunctions();

        //Main Function of Test1 Class
        public void Main()
        {
            SldWorks swApp = null;
            ModelDoc2 swModel = null;
            PartDoc swPart = null;
            SelectionMgr swSelectionMgr = null;
            View swView = null;
            Entity swEntity1 = null;
            Entity swEntity2 = null;
            DrawingDoc drawing = null;

            bool edge1Status = false;
            bool edge2Status = false;
            bool edgeStatus = false;
            object dimStatus = null;
            string edgeName1 = null;
            string edgeName2 = null;
            

            try
            {

                //Opening New Instance of Solidworks Application
                swApp = new SldWorks();

                // Making Solidworks Application Visible to User
                swApp.Visible = true;

                // Get reference to the active model document
                swModel = (ModelDoc2)swApp.ActiveDoc;

                //Checks Solidworks App is Opened with Part Type
                if (swModel != null && swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
                {

                    //Clear the Existing Selection Filters
                    cmFunc.ClearSelFilters(swApp);
                    
                    // Set Selection Filter for only Edges
                    swApp.SetSelectionFilter(1, true);

                    // Msg for User
                    MessageBox.Show("Select only Two Edges, which are Suitable for Linear Dimension to Process");

                    //Cast ModelDoc2 to PartDoc
                    swPart = (PartDoc)swModel;

                    // Get a reference to the selection manager
                    swSelectionMgr = (SelectionMgr)swModel.SelectionManager;

                    // Get the first edge
                    swEntity1 = GetEdgeFromPart(swModel);


                    // Get the second edge
                    while (swEntity2 == null)
                    {
                        
                        swEntity2 = GetEdgeFromPart(swModel,"Please select Second Edge (one Edge only), Then give OK", "Edge2");

                        // Ensure Edge1 is not equal Edge2
                        if (swEntity2!=null && swEntity1 != swEntity2)
                        {
                           
                            //Get associated Drawing
                            while (swModel != null && swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                            {

                                // Set Refernce to opened Drawing
                                swModel = (ModelDoc2)swApp.ActiveDoc;

                                // Ensure opened document is drawing
                                if (swModel != null && swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                                {
                                    
                                    // Get a reference to the selection manager
                                    swSelectionMgr = (SelectionMgr)swModel.SelectionManager;

                                    // To Get associated View
                                    while (swView == null)
                                    {
                                        // Get the Selected view
                                        swView = (View)swSelectionMgr.GetSelectedObject6(1, -1);

                                        // Ensure Drawing view is selected
                                        if (swSelectionMgr.GetSelectedObjectType(1) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                                        {
                                            
                                            //Get the part from selected Drawing view
                                            swPart = (PartDoc)swView.ReferencedDocument;

                                            //Show & Select edge1 in Drawing view
                                            edge1Status = swView.SelectEntity(swEntity1, false);

                                            //Show & Select edge2 in Drawing view
                                            edge2Status = swView.SelectEntity(swEntity2, true);

                                            //Add Dimensions for above selected Edges
                                            dimStatus = swModel.AddDimension(0.0, 0.0, 0.0);

                                            //Ensure to clear any other selection Made eariler
                                            swModel.ClearSelection2(true);

                                            //To Highlight the selected Edges in Drawing, Again i am Calling the SelectEntity.
                                            //Show edge1 in Drawing view
                                            edge1Status = swView.SelectEntity(swEntity1, false);

                                            //Show edge2 in Drawing view
                                            edge2Status = swView.SelectEntity(swEntity2, true);


                                        }
                                        else
                                        {
                                            //Message for user with Cancel option, while in Loop.
                                            cmFunc.MsgBoxWithCancelBt("Select an Drawing view, Then give OK", "Drawing view");

                                        }
                                    }
                                }
                                else
                                {
                                    //Message for user with Cancel option, while in Loop.
                                    cmFunc.MsgBoxWithCancelBt("Open Drawing file in Solidworks, Then give OK", "Drawing File");

                                }


                            }

                           

                        }
                        else
                        {
                            //Message for user with Cancel option, while in Loop.
                            cmFunc.MsgBoxWithCancelBt("Please select Diffrent Edge. Edge1 & Edge2 should not be unique, Then give OK", "Select Diffrent Edge");
                            swEntity2 = null;
                        }


                    }

                    // Checks Both edge has been shown in Drawing
                    edgeStatus = (edge1Status && edge2Status) ? true : false;


                    #region Final Status & message to user
                    if (dimStatus == null&& edgeStatus)
                    {
                        // Msg for User
                        MessageBox.Show("Edge has been shown in Drawing. But Edges Selected in Part is not suitable for Linear Dimension. Please do the process again, by Pressing Run Button Again");

                    }

                    else if(dimStatus != null && edgeStatus == false)
                    {
                        // Msg for User
                        MessageBox.Show("Linear Dimension Placed.But Due to some reason, Edge couldn't able to show in drawing. Please do the process again, by Pressing Run Button Again");

                    }
                    else if (dimStatus == null && edgeStatus == false)
                    {
                        // Msg for User
                        MessageBox.Show("Edges & Linear Dimesnion couldn't able process in selected drawing View. Please do the process again, by Pressing Run Button Again");

                    }
                    else
                    {
                        // Msg for User
                        MessageBox.Show("Edges has been shown & Linear Dimesnion has been placed in selected drawing View");
                    }

                    #endregion


                }
                else
                {
                    // Msg for User
                    MessageBox.Show("Open Part file in Solidworks & Press Run Button Again");

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
               
                if (swModel != null) Marshal.ReleaseComObject(swModel);
                if (swApp != null) Marshal.ReleaseComObject(swApp);
                if (swView != null) Marshal.ReleaseComObject(swView);
                if (swEntity1 != null) Marshal.ReleaseComObject(swEntity1);
                if (swEntity2 != null) Marshal.ReleaseComObject(swEntity2);
                if (swSelectionMgr != null) Marshal.ReleaseComObject(swSelectionMgr);
                if (swPart != null) Marshal.ReleaseComObject(swPart);
                if (drawing != null) Marshal.ReleaseComObject(drawing);



            }

        }

        //Function to get Edge from selected Part
        public Entity GetEdgeFromPart(ModelDoc2 swModel ,string msgBoxContent = "Please select First Edge (one Edge only), Then give OK", string msgBoxHeader = "Edge1")
        {
           
            PartDoc swPart = null;
            Edge edge = null;
            SelectionMgr swSelectionMgr = null;
            Entity swEntity = null;
            //bool boolStatus = false;
            int edgeCount = 0;
            //string currentEdgeName;

            
            //Cast ModelDoc2 to PartDoc
            swPart = (PartDoc)swModel;

            // Get a reference to the selection manager
            swSelectionMgr = (SelectionMgr)swModel.SelectionManager;

                    // Get the edge
                    
                    while (edge == null)
                    {
                        // Get the selected object count
                        edgeCount = swSelectionMgr.GetSelectedObjectCount2(-1);

                        //assign selected object to entity
                        swEntity = (Entity)swSelectionMgr.GetSelectedObject6(edgeCount, -1);
                        
                        //Make sure selected was edge and only one edge was selected
                        if (swEntity is Edge && edgeCount == 1)
                        {
                            //Get Edge to Close the While Loop.
                            edge = (Edge)swEntity;

                            //Clear the selection which was selected eariler
                            swModel.ClearSelection2(true);

                        }
                        else

                        {
                            //Message for user with Cancel option, while in Loop.
                            cmFunc.MsgBoxWithCancelBt(msgBoxContent, msgBoxHeader);


                        }



                    }

            return swEntity;



        }


    }

    public class Test2
    {
        CommonFunctions cmFunc = new CommonFunctions();

        public void Main(string txtPath)
        {
            SldWorks swApp = null;
            ModelDoc2 swModel = null;
            SelectionMgr swSelMgr = null;
            AttributeDef swAttDef = null;
            Attribute swAtt = null;
            Parameter swParam = null;
            Face2 swFace = null;
            Feature swFeat = null;
            Entity swEnt = null;

            DialogResult dialogResult;

            string attName = "FACE_ID";
            string attValue = "FACE_1";
            string loopCount;
            int faceCount = 2;
            bool boolStatus;

            string faceID = "";
            string persisID;

            try
            {
                swApp = new SldWorks();
                swApp.Visible = true;
                swModel = (ModelDoc2)swApp.ActiveDoc;

                //Checks Solidworks App is Opened with Part or Assembly Type only
                if (swModel != null && (swModel.GetType() == (int)swDocumentTypes_e.swDocPART|| swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY))
                {

                    //Clear all existing selection filters
                    cmFunc.ClearSelFilters(swApp);


                    // Set Selection Filter for only Faces
                    swApp.SetSelectionFilter(2, true);

                    // Msg for User
                    MessageBox.Show("Select any one face, which you need to get Persistence ID of it");


                    swSelMgr = (SelectionMgr)swModel.SelectionManager;

                    
                    // To get Face
                    while (swFace == null)
                    {
                        faceCount = swSelMgr.GetSelectedObjectCount2(-1);

                        // Get face to be selected by user
                        swFace = (Face2)swSelMgr.GetSelectedObject6(faceCount, -1);

                        swEnt = (Entity)swFace;

                        // Ensure Selection was face & only one Face
                        if (swFace is Face2 && faceCount == 1)
                        {
                            // Get the PersistenceID of Selected Face
                            persisID = GetPersistenceID(swModel,swFace);

                            //Clear all existing selection filters
                            cmFunc.ClearSelFilters(swApp);

                            #region Add Attribute to selected Faces

                            // Select the Attribute from design Tree
                            boolStatus = swModel.Extension.SelectByID2(attName, "ATTRIBUTE", 0, 0, 0, false, 0, null, 0);

                            //Ensure Attribute is Selected
                            if (boolStatus)
                            {
                                //Get Attribute as Feature from Design Tree
                                swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);

                                //Cast Feature to Attribute
                                swAtt = (Attribute)swFeat.GetSpecificFeature2();

                                //Delete if Already Attribute is Created.
                                swAtt.Delete(true);

                                //Make Force Rebuild
                                swModel.ForceRebuild3(true);

                            }
                                   

                            //Define Attribute
                            swAttDef = (AttributeDef)swApp.DefineAttribute(attName);

                            //Add Parameter
                            swAttDef.AddParameter(faceID, (int)swParamType_e.swParamTypeString, 0.0, 0);

                            //Register the Added Parameters
                            swAttDef.Register();

                            //Create New Instance of Attribut ID in Design Tree
                            swAtt = (Attribute)swAttDef.CreateInstance5(swModel, swEnt, attName, 0, 2);

                            //Force Rebuild
                            swModel.ForceRebuild3(true);

                            //Get the Created Artibute
                            swParam = (Parameter)swAtt.GetParameter(faceID);

                            //Declare the value for created Parameters
                            swParam.SetStringValue2(persisID, (int)swInConfigurationOpts_e.swAllConfiguration, "");

                            // Calling below function to Save the Presistance ID to a Text file
                            SaveDataToTextFile(txtPath, faceID + " = " + persisID);

                            #endregion

                            //Clear all existing selection filters
                            cmFunc.ClearSelFilters(swApp);

                            //Clear all the Selected Entity
                            swModel.ClearSelection2(true);


                        }
                        else
                        {
                            //Message for user with cancel Button
                            cmFunc.MsgBoxWithCancelBt("Please select an Face to get its Persistence ID, Then give OK", "Select any one Face");
                            
                        }


                    }

                       MessageBox.Show("Attribute are Created and Persistence ID Saved in your Intial Selection of Text File");


                }
                else
                {
                    // Msg for User
                    MessageBox.Show("Open Part file in Solidworks & Give Run Again");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                
                if (swParam != null) Marshal.ReleaseComObject(swParam);
                if (swAtt != null) Marshal.ReleaseComObject(swAtt);
                if (swAttDef != null) Marshal.ReleaseComObject(swAttDef);
                if (swFace != null) Marshal.ReleaseComObject(swFace);
                if (swSelMgr != null) Marshal.ReleaseComObject(swSelMgr);
                if (swModel != null) Marshal.ReleaseComObject(swModel);
                if (swApp != null) Marshal.ReleaseComObject(swApp);
                if (swFace != null) Marshal.ReleaseComObject(swFace);
                if (swEnt != null) Marshal.ReleaseComObject(swEnt);


            }
        }

        //Get Persistence ID from Face
        public string GetPersistenceID(ModelDoc2 swModel,Face2 swFace)
        {
            string base64String;
            byte[] vPIDarr;
            
            
            //Gets Persistence ID 
            vPIDarr = (byte[])swModel.Extension.GetPersistReference3(swFace);

            //Convert Object data to string
            base64String = Convert.ToBase64String(vPIDarr);

            return base64String;


        }

        // Function to Save the Presistence ID to Text file
        public void SaveDataToTextFile(string filePath, string TextOutput)
        {
            
            //Ensure Safe Handling of Text File
            using (var writer = new StreamWriter(filePath))
            {
                // Write to text File
                writer.WriteLine(TextOutput);
                // Close the Text file
                writer.Close();

            }
               
        }


    }

    public class Test3
    {
        CommonFunctions cmFunc = new CommonFunctions();

        public void Main()
        {

            SldWorks swApp = null;
            ModelDoc2 swModel = null;
            SelectionMgr swSelMgr = null;
            PartDoc swPart = null;
            Component2 swComponent = null;
            AssemblyDoc swAssembly = null;
            Face2 swFace = null;

            bool boolStatus = false;
            int compCount;
            int faceCount;
            int errorCount = 0;
            
            string assmName;
            string partName = null;
            string partNamewithInstance;

            
            try
            {
                Entity swInputEntity;
                Entity swOutputEntity;
                int Errors = 0;


                swApp = new SldWorks();
                swApp.Visible = true;
                swModel = (ModelDoc2)swApp.ActiveDoc;
                

                //Checks Solidworks App is Opened with Part Type
                if (swModel != null && swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    //Casting Modeldoc2 to AssemblyDoc
                    swAssembly = (AssemblyDoc)swModel;

                    //Clear Existing selection Made by User
                    swModel.ClearSelection2(true);

                    //Giving note to user
                    MessageBox.Show("Open Part file from Assembly, Then Select any one face from the part");

                    //Get the Assembly Name from Opened Document.
                    assmName = swModel.GetTitle();

                    //Ensure Part was opened
                    while (partName == null) 
                    { 
                        //Assign currently opened part file as Active Doc
                        swModel = (ModelDoc2)swApp.ActiveDoc;

                        //Ensure currently opened Document is Part Document
                        if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
                        {
                            //Casting Modeldoc2 to PartDoc
                            swPart = (PartDoc)swModel;

                            //Get the Part Name from Opened Document & Removing its extension.
                            //Because get component by Name Method dosen't Require extension
                            partName = swModel.GetTitle().Replace(".SLDPRT","");

                            //Clear the Existing Selection Filters
                            cmFunc.ClearSelFilters(swApp);

                            // Set Selection Filter for only Faces
                            swApp.SetSelectionFilter(2, true);

                            // Calling selection Manager to get selection from User
                            swSelMgr = (SelectionMgr)swModel.SelectionManager;

                            //Got face has Input from User
                            swInputEntity = (Entity)swSelMgr.GetSelectedObject6(1, -1);

                            //Until face was selected Loop will go on
                            while (swFace == null)
                            {
                                //Getting count of Face
                                faceCount = swSelMgr.GetSelectedObjectCount2(-1);

                                //Getting selected object as Face with Selected count
                                swFace = (Face2)swSelMgr.GetSelectedObject6(faceCount, -1);

                                //Casting face into Entity 
                                swInputEntity = (Entity)swFace;

                                //Ensure Selected object is face and only one was selected
                                if (swFace is Face2 && faceCount == 1)
                                {
                                    //Clear the Existing Selection Filters
                                    cmFunc.ClearSelFilters(swApp);

                                    //Clear the selected Entity from Part
                                    swModel.ClearSelection2(true);

                                    //Switching Part Document window to Assembly Window. (Already Opened window at Program Start)
                                    swModel = (ModelDoc2)swApp.ActivateDoc3(assmName, false, (int)swRebuildOnActivation_e.swRebuildActiveDoc, ref Errors);
                                    
                                    //Set Opened Assembly Document as Active Doc
                                    swModel = (ModelDoc2)swApp.ActiveDoc;

                                    //Casting Modeldoc2 to AssemblyDoc
                                    swAssembly = (AssemblyDoc)swModel;

                                    //Get Number of Components(Part) in Top Level of Assembly
                                    compCount = swAssembly.GetComponentCount(true);

                                    //Run Loop for Number of Components in Assembly.
                                    //So we will ensure, we have selected all applicable Components in Assembly
                                    for(int i = 0; i < compCount; i++)
                                    {
                                        //Ensure Loop Start with Instance Number of 1
                                        int instanStartNum = i + 1;

                                        //Set Selected Part Name with Instance number
                                        partNamewithInstance = partName+"-"+ instanStartNum.ToString();

                                        //Check the part is avialable in Assembly with instance Number.
                                        swComponent = (Component2)swAssembly.GetComponentByName(partNamewithInstance);

                                        //Ensure Part is Available
                                        if (swComponent != null)
                                        {
                                            //Get the Corresponding Entity from above selected Part
                                            swOutputEntity = (Entity)swComponent.GetCorrespondingEntity(swInputEntity);

                                            //Select the above Entity in Graphic window, Without clearing previous Selection
                                            boolStatus = swOutputEntity.Select4(true, null);

                                        }
                                        


                                    }
                                    



                                }
                                else
                                {
                                    //Message to user with Cancel Option
                                    cmFunc.MsgBoxWithCancelBt("Select any one face from the part", "Select Face");

                                }


                            }





                        }

                        else
                        {
                            //Message to user with Cancel Option
                            cmFunc.MsgBoxWithCancelBt("Open Part file from Assembly", "Open Part File");

                        }



                    }





                }
                else
                {

                    MessageBox.Show("Open Solidworks Assembly with Instance Parts & Press Run Button Again");
                    errorCount =1;

                }

                if (boolStatus)
                {

                    MessageBox.Show("Faces of all instances has been Shown in Assembly");

                }
                else if (!boolStatus && errorCount == 0)
                {
                    MessageBox.Show("Couldn't able to Show all the Faces in Assembly, Hit the Run button Again");
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                

                if (swSelMgr != null) Marshal.ReleaseComObject(swSelMgr);
                if (swModel != null) Marshal.ReleaseComObject(swModel);
                if (swApp != null) Marshal.ReleaseComObject(swApp);
                if(swPart != null) Marshal.ReleaseComObject(swPart);
                if (swComponent != null) Marshal.ReleaseComObject(swComponent);
                if (swAssembly != null) Marshal.ReleaseComObject(swAssembly);
                if (swFace != null) Marshal.ReleaseComObject(swFace);

            }


        }

    }

    public class Test4
    {

        public void Main()
        {
            SldWorks swApp = null;
            ModelDoc2 swModel = null;
            ModelView swModelView = null;
            string viewNameNew = "New View 1";
          


            try
            {
                swApp = new SldWorks();
                swApp.Visible = true;
                swModel = (ModelDoc2)swApp.ActiveDoc;

                //Checks Solidworks App is Opened with Part or Assembly Type Only 
                if (swModel != null && (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY|| swModel.GetType() == (int)swDocumentTypes_e.swDocPART))
                {
                    //Set Model view as Active View
                    swModelView = swModel.IActiveView;

                    //Rotate the Model to Create Custom View. Rotate about centre in Y angle.
                    swModelView.RotateAboutCenter(0, 10);

                    //Create New View with Custom Name
                    swModel.NameView(viewNameNew);

                    //Show the Created new view in Model, Ensure Only mentioned view has been Shown
                    swModel.ShowNamedView2(viewNameNew, -1);

                    MessageBox.Show("The new view was created with Name of " + viewNameNew + " ,The View Rotaion angle was Hard Coded!!");
                }
                else
                {
                    MessageBox.Show("Open Solidworks Assembly or Part to create Custom View & Press Run Button Again");

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {


                if (swModel != null) Marshal.ReleaseComObject(swModel);
                if (swApp != null) Marshal.ReleaseComObject(swApp);
                if (swModelView != null) Marshal.ReleaseComObject(swModelView);


            }


        }
    }

    public class Test5
    {

        public void Main(string filePath)
        {
            SldWorks swApp = null;
            ModelDoc2 swModel = null;
            AssemblyDoc swAssembly = null;
            InterferenceDetectionMgr pIntMgr = null;
            object[] vInts = null;
            long i = 0;
            long j = 0;
            Interference interference = null;
            object vIntComps = null;
            object[] vComps = null;
            Component2 swComponent = null;
            //double vol = 0;
            object vTrans = null;
            bool ret = false;
            try
            {
                swApp = new SldWorks();
                swApp.Visible = true;
                swModel = (ModelDoc2)swApp.ActiveDoc;

                //Ensure Assembly Document is Opened
                if (swModel != null && swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    //Message to User
                    MessageBox.Show("If Interfence detected between Parts in Assembly, Then Result will be Saved in your Intial Selection of Text File");
                    
                    //Casting ModelDoc2 to Assembly Doc
                    swAssembly = (AssemblyDoc)swModel;

                    //Open the Interference Detection pane
                    swAssembly.ToolsCheckInterference();
                    pIntMgr = swAssembly.InterferenceDetectionManager;


                    //Specify the interference detection settings and options
                    pIntMgr.TreatCoincidenceAsInterference = false;
                    pIntMgr.TreatSubAssembliesAsComponents = false;
                    pIntMgr.IncludeMultibodyPartInterferences = true;
                    pIntMgr.MakeInterferingPartsTransparent = true;
                    pIntMgr.CreateFastenersFolder = false;
                    pIntMgr.IgnoreHiddenBodies = false;
                    pIntMgr.ShowIgnoredInterferences = false;
                    pIntMgr.UseTransform = false;


                    //Specify how to display non-interfering components
                    pIntMgr.NonInterferingComponentDisplay = (int)swNonInterferingComponentDisplay_e.swNonInterferingComponentDisplay_Wireframe;


                    //Run interference detection
                    vInts = (object[])pIntMgr.GetInterferences();
                    if (vInts != null)
                    {
                        //Safe Method operating Text file
                        using (var writer = new StreamWriter(filePath))
                        {
                            //Write Number of Interfering Coponent Count
                            writer.WriteLine("Number of interferences: " + pIntMgr.GetInterferenceCount());

                            //Get interfering components and transforms
                            ret = pIntMgr.GetComponentsAndTransforms(out vIntComps, out vTrans);

                            //Get interference Group information of all Components, & Loop Through all.
                            for (i = 0; i <= vInts.GetUpperBound(0); i++)
                            {
                                //Interference Group
                                writer.WriteLine("Interference " + (i + 1));

                                //Casting Interfered Componets to Interefnce
                                interference = (Interference)vInts[i];

                                //Write Number of Compontent has been interfering in Current Group
                                writer.WriteLine("  Number of components in this interference: " + interference.GetComponentCount());
                               
                                //Get all Components of Current Group
                                vComps = (object[])interference.Components;

                                //Loop through all the Components in the Group.
                                for (j = 0; j <= vComps.GetUpperBound(0); j++)
                                {
                                    //Cast interfernce Component to Component2
                                    swComponent = (Component2)vComps[j];

                                    //Write the Name of all Interference Component in Current Group.
                                    writer.WriteLine("   " + swComponent.Name2);

                                }


                            }

                            // Close the Text file
                            writer.Close();

                            MessageBox.Show("Parts are Interfering in Assembly, Results are Saved in your Intial Selection of Text File");

                        }

                    }
                    else
                    {
                        MessageBox.Show("Parts are Not Interfering in Assembly");

                    }


                    //Stop interference detection and close Inteference Detection pane
                    pIntMgr.Done();

                }
                else
                {
                    MessageBox.Show("Open Solidworks Assembly to check Interfence Between Parts & Press Run Button Again");


                }



            }


            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {


                if (swModel != null) Marshal.ReleaseComObject(swModel);
                if (swApp != null) Marshal.ReleaseComObject(swApp);
                if (swComponent != null) Marshal.ReleaseComObject(swComponent);
                if (swAssembly != null) Marshal.ReleaseComObject(swAssembly);
                if (pIntMgr != null) Marshal.ReleaseComObject(pIntMgr);
                if (interference != null) Marshal.ReleaseComObject(interference);

            }



        }
      

    }



}


















