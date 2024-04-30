using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using System.Globalization;
using System.Security.Cryptography;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace CTI_Utilities
{
    public partial class CTI_Utilties
    {

        private void CTI_Utilties_Load(object sender, RibbonUIEventArgs e)
        {
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.Button1_Click);

            Globals.ThisAddIn.Application.DocumentCreated += new EApplication_DocumentCreatedEventHandler(NewDocument_Created);
            Globals.ThisAddIn.Application.DocumentOpened += new EApplication_DocumentOpenedEventHandler(NewDocument_Opened);

            

        }
        private void NewDocument_Created(Document doc)
        {
            System.Diagnostics.Debug.WriteLine("New Document Created");
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Visio.Shape entryShape;
            Visio.Pages pages = document.Pages;
            Visio.Page page = pages[1];
            RevEntry revEntry = new RevEntry();
            int intRevX2;


            for (intRevX2 = page.Shapes.Count; intRevX2 >= 1; intRevX2--)
            {
                entryShape = page.Shapes[intRevX2];
                bool b = entryShape.Name.Contains("Revision");
                if (b)
                {
                    revEntry.DeleteCoverEntries(entryShape);

                }
            }

            revEntry.DropFullEntry();
            Globals.ThisAddIn.Application.ActiveDocument.DocumentOpened += new EDocument_DocumentOpenedEventHandler(NewDocument_Opened);
            Globals.ThisAddIn.Application.ActiveDocument.PageAdded += new EDocument_PageAddedEventHandler(NewPage_Added);
        }
        private void NewDocument_Opened(Document doc)
        {
            System.Diagnostics.Debug.WriteLine("New Document Opened");
            Globals.ThisAddIn.Application.ActiveDocument.PageAdded += new EDocument_PageAddedEventHandler(NewPage_Added);
            Globals.ThisAddIn.Application.DocumentOpened -= new EApplication_DocumentOpenedEventHandler(NewDocument_Opened);
            Globals.ThisAddIn.Application.BeforeDocumentClose += new EApplication_BeforeDocumentCloseEventHandler(NewDocument_Closed);

            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            const string CUST_PROP_PREFIX = "Prop.";
            Visio.Page lastPage = Globals.ThisAddIn.Application.ActiveDocument.Pages[Globals.ThisAddIn.Application.ActiveDocument.Pages.Count];
            Visio.Cell customPropertyCell;
            string strDesignedBy = "DesignedBy";
            string strApprovedBy = "ApprovedBy";
            string strDrawnBy = "DrawnBy";

            if (document.Company != null)
            {
                editBox2.Text = document.Company.ToString();
            }
            if (document.Title != null)
            {
                editBox3.Text = document.Title.ToString();
            }
            if (document.Subject != null)
            {
                editBox4.Text = document.Subject.ToString();
            }
            if (document.Creator != null)
            {
                editBox10.Text = document.Creator.ToString();
            }
            if (editBox8 != null && editBox9 != null && editBox10 != null)
            {
                try
                {
                    foreach (Visio.Shape revShape in lastPage.Shapes)
                    {
                        bool b = revShape.Name.Contains("TitleBlockNames");
                        if (b)
                        {
                            customPropertyCell = revShape.get_CellsU(CUST_PROP_PREFIX + strApprovedBy);
                            editBox8.Text = customPropertyCell.Formula.ToString().Trim('\"');
                            customPropertyCell = revShape.get_CellsU(CUST_PROP_PREFIX + strDesignedBy);
                            editBox9.Text = customPropertyCell.Formula.ToString().Trim('\"');
                            customPropertyCell = revShape.get_CellsU(CUST_PROP_PREFIX + strDrawnBy);
                            editBox10.Text = customPropertyCell.Formula.ToString().Trim('\"');
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {

                }
            }
        }
        private void NewDocument_Closed(Document doc)
        {
            System.Diagnostics.Debug.WriteLine("New Document Closed");
            Globals.ThisAddIn.Application.ActiveDocument.PageAdded -= new EDocument_PageAddedEventHandler(NewPage_Added);
            Globals.ThisAddIn.Application.DocumentOpened += new EApplication_DocumentOpenedEventHandler(NewDocument_Opened);
            Globals.ThisAddIn.Application.BeforeDocumentClose -= new EApplication_BeforeDocumentCloseEventHandler(NewDocument_Closed);
        }
        private void NewPage_Added(Visio.Page Page)
        {
            
            System.Diagnostics.Debug.WriteLine("New Page Added");

            Visio.Documents visioDocuments = Globals.ThisAddIn.Application.Documents;
            Visio.Document stencil;
            Visio.Master masterInStencil;
            Visio.Shape drawingTitle, pageShape;
            pageShape = Page.PageSheet;
            AddingACustomProperty newCustomProp = new AddingACustomProperty();

            

            try
            {
                stencil = visioDocuments["TitlePage Stencil.vssx"];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // The stencil is not in the collection; open it as a 
                // docked stencil.
                stencil = visioDocuments.OpenEx("TitlePage Stencil.vssx",
                    (short)Microsoft.Office.Interop.Visio.
                    VisOpenSaveArgs.visOpenDocked);
            }

            foreach(Visio.Shape checkShapes in Page.Shapes)
            {
                if(checkShapes.Name.Contains("PageTitle"))
                    {
                        stencil.Close();
                        return;
                    }
            }

            masterInStencil = stencil.Masters.get_ItemU("PageTitle");
            drawingTitle = masterInStencil.Shapes[1];
            System.Diagnostics.Debug.WriteLine("Dropping Shape" + drawingTitle.Name);
            Page.Drop(drawingTitle, 32.625, 2.25);
            System.Diagnostics.Debug.WriteLine("Shape Dropped");
            newCustomProp.AddCustomProperty(pageShape, "RoomName", "Room", "Room Name",VisCellVals.visPropTypeString, "@", "", false, false, "1");
            Page.Application.ActiveWindow.DeselectAll();
            stencil.Close();
            return;
        }

        //CABLE LABEL AUTONUMBER
        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            int startingNum,incrementNum;

            if (this.editBox1.Text == "")
            {
                startingNum = 101;
            }
            else
            {
                startingNum = int.Parse(this.editBox1.Text);
            }
            if (this.checkBox2.Checked == false)
            {
                incrementNum = 1;
            }
            else
            {
                incrementNum = 100;
            }
            AutoNumbering("Cable_Num", startingNum, incrementNum, true);
        }

        //AUTO NUMBERING '#' SHAPE DATA
        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            int x, Count;
            string tempToString, tempFromString;
            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Document document = visioApplication.ActiveDocument;
            Visio.Shape visShape;
            const string CUST_PROP_PREFIX = "Prop.";


            Visio.Pages visioPages = document.Pages;
            Visio.Page page = visioPages[1];

            var myPageCol = new List<Visio.Page>();
            var myShapeCol = new List<Visio.Shape>();

            if (this.editBox5.Text == "")
            {
                Count = 1;
            }
            else
            {
                Count = int.Parse(this.editBox5.Text);
            }

            foreach (Visio.Page PageToIndex in document.Pages)
            {
                if (checkBox1.Checked == true)
                {
                    myPageCol.Add(PageToIndex);
                }
                else
                {
                    myPageCol.Add(visioApplication.ActivePage);
                }
            }


            foreach (Visio.Page PageToIndex in myPageCol)
            {
                if (PageToIndex.Background == page.Background)
                {
                    int intShapeCt, intX;
                    intShapeCt = PageToIndex.Shapes.Count;
                    for (intX = 1; intX <= intShapeCt; intX++)
                    {
                        visShape = PageToIndex.Shapes[intX];
                        bool b = visShape.Name.Contains("Pulled Cable V3");
                        if (b)
                        {
                            myShapeCol.Add(visShape);
                        }
                    }
                }
                else
                {

                }
            }
            while (myShapeCol.Count > 0)
            {

                Visio.Cell tempPropertyCell, compareToCell, compareFromCell, compareHashtagCell;
                string compareToString, compareFromString;
                Visio.Shape tempPropertyShape = myShapeCol[0];

                // Get the Cell object. Note the addition of "Prop." to the
                // name given to the cell.
                tempPropertyCell = tempPropertyShape.get_CellsU(CUST_PROP_PREFIX + "Cable_To");
                tempToString = tempPropertyCell.Formula;
                tempPropertyCell = tempPropertyShape.get_CellsU(CUST_PROP_PREFIX + "Cable_From");
                tempFromString = tempPropertyCell.Formula;
                for (x = myShapeCol.Count - 1; x >= 0; x--)
                {
                    Visio.Shape comparePropertyShape = myShapeCol[x];
                    compareToCell = comparePropertyShape.get_CellsU(CUST_PROP_PREFIX + "Cable_To");
                    compareFromCell = comparePropertyShape.get_CellsU(CUST_PROP_PREFIX + "Cable_From");
                    compareHashtagCell = comparePropertyShape.get_CellsU(CUST_PROP_PREFIX + "Cable_HashTag");
                    compareToString = compareToCell.Formula;
                    compareFromString = compareFromCell.Formula;
                    if (compareToString == tempToString && compareFromString == tempFromString)
                    {
                        compareHashtagCell.FormulaU = "\"" + Count.ToString("00") + "\""; 
                        myShapeCol.RemoveAt(x);
                    }
                }
                Count++;
            }
        }

        //AUTO NUMBERING "I/O" SHAPE DATA
        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {
            int startingNum;
            if (this.editBox6.Text == "")
            {
                startingNum = 1;
            }
            else
            {
                startingNum = int.Parse(this.editBox6.Text);
            }
            AutoNumbering("IO_NUM", startingNum, 1, false);
        }
        
        //AUTO NUMBERING "DEVICE NUMBER" SHAPE DATA
        private void Button4_Click(object sender, RibbonControlEventArgs e)
        {
            int startingNum;
            if (this.editBox7.Text == "")
            {
                startingNum = 1;
            }
            else
            {
                startingNum = int.Parse(this.editBox7.Text);
            }
            AutoNumbering("DeviceNum", startingNum, 1, true);
        }

        //RESET CABLE LABEL AUTONUMBER FIELD
        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBox1.Text = "";
            return;
        }

        //RESET "#" AUTONUMBER FIELD
        private void Button6_Click(object sender, RibbonControlEventArgs e)
        {
            editBox5.Text = "";
        }

        //RESET "I/O" AUTONUMBER FIELD
        private void Button7_Click(object sender, RibbonControlEventArgs e)
        {
            editBox6.Text = "";
        }
        
        //RESET "DEVICE NUMBER" AUTONUMBER FIELD
        private void Button8_Click(object sender, RibbonControlEventArgs e)
        {
            
            editBox7.Text = "";
        }
        
        //TABLE OF CONTENTS UPDATE
        private void Button9_Click(object sender, RibbonControlEventArgs e)
        {
            TOCEntry(true);
        }

        //NEW MAJOR REVISION ENTRY
        private void Button10_Click(object sender, RibbonControlEventArgs e)
        {
            this.splitButton2.Click -= new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.Button11_Click);
            this.splitButton2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.Button10_Click);
            this.splitButton2.Label = "New Major Revision";
            this.splitButton2.OfficeImageId = "MasterPageAddNew";
            Visio.Documents visioDocuments = Globals.ThisAddIn.Application.Documents;
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Visio.Shape entryShape;
            Visio.Pages pages = document.Pages;
            Visio.Page page = pages[1];
            RevEntry revEntry = new RevEntry();
            int intRevX2;


            for (intRevX2 = page.Shapes.Count; intRevX2 >= 1; intRevX2--)
            {
                entryShape = page.Shapes[intRevX2];
                bool b = entryShape.Name.Contains("Revision");
                if (b)
                {
                    revEntry.boolMajorRev = true;
                    revEntry.revType = dropDown1.SelectedItem.ToString();
                    revEntry.SetShape(entryShape);
                    revEntry.DropFullEntry();
                    RevLevel_Label.Text = revEntry.RevNum.ToString();
                    return;
                }
            }
        }

        //NEW MINOR REVISION ENTRY
        private void Button11_Click(object sender, RibbonControlEventArgs e)
        {
            this.splitButton2.Click -= new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.Button10_Click);
            this.splitButton2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.Button11_Click);
            this.splitButton2.Label = "New Minor Revision";
            this.splitButton2.OfficeImageId = "NewMaster";

            Visio.Documents visioDocuments = Globals.ThisAddIn.Application.Documents;
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Visio.Shape entryShape;
            Visio.Pages pages = document.Pages;
            Visio.Page page = pages[1];
            RevEntry revEntry = new RevEntry();
            int intRevX2;


            for (intRevX2 = page.Shapes.Count; intRevX2 >= 1; intRevX2--)
            {
                entryShape = page.Shapes[intRevX2];
                bool b = entryShape.Name.Contains("Revision");
                if (b)
                {
                    revEntry.boolMajorRev = false;
                    revEntry.revType = dropDown1.SelectedItem.ToString();
                    revEntry.SetShape(entryShape);
                    revEntry.DropFullEntry();
                    RevLevel_Label.Text = revEntry.RevNum.ToString();
                    return;
                }
            }
        }

        //SET HOSTNAMES
        private void Button12_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Document document = visioApplication.ActiveDocument;
            Visio.Pages visioPages = document.Pages;
            var myShapeCol = new List<Visio.Shape>();
            const string CUST_PROP_PREFIX = "Prop.";
            Visio.Cell tempPropertyCell;
            string roomString;


            foreach (Visio.Page visioPage in visioPages)
            {
                int intShapeCt, intX;
                Visio.Shape visShape;
                Visio.Shape pageSheet;
                intShapeCt = visioPage.Shapes.Count;
                for (intX = 1; intX <= intShapeCt; intX++)
                {
                    visShape = visioPage.Shapes[intX];
                    bool b = visShape.Name.Contains("Equipment Stencil");
                    if (b)
                    {
                        myShapeCol.Add(visShape);
                        pageSheet = visioPage.PageSheet;
                        tempPropertyCell = pageSheet.get_CellsU(CUST_PROP_PREFIX + "RoomName");
                        roomString = tempPropertyCell.Formula;
                        roomString = roomString.Replace(" ", "");
                        HostNaming(visShape, roomString);
                    }
                }
            }
        }

        //SET DEVICE PAGE/DEVICE SCHEDULE
        private void Button13_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Document document = visioApplication.ActiveDocument;
            const string CUST_PROP_PREFIX = "Prop.";
            AddingACustomProperty newCustomProp = new AddingACustomProperty();
            Visio.Shape pageShape;

            Visio.Pages visioPages = document.Pages;
            Visio.Page page = visioPages[1];
            pageShape = page.PageSheet;

            Visio.Cell PageNameCell;
            Visio.Cell tempPropertyCell;
            string roomString;

            foreach (Visio.Page PageToIndex in document.Pages)
            {
                if (PageToIndex.Background == page.Background)
                {
                    Visio.Page IndexingPage;
                    IndexingPage = PageToIndex;
                    int intShapeCt;
                    intShapeCt = IndexingPage.Shapes.Count;
                    Visio.Shape pageSheet;


                    foreach (Visio.Shape visShapes in document.Pages[PageToIndex.Index].Shapes)
                    {
                        bool b = visShapes.NameU.Contains("Equipment Stencil");
                        bool c = visShapes.NameU.Contains("Mount Stencil");
                        bool d = visShapes.NameU.Contains("Speaker Stencil");
                        bool f = visShapes.NameU.Contains("Pulled Cable V3");
                        if (b || c || d || f)
                        {
                            try
                            {
                                pageSheet = document.Pages[PageToIndex.Index].PageSheet;
                                tempPropertyCell = pageSheet.get_CellsU(CUST_PROP_PREFIX + "RoomName");
                                roomString = tempPropertyCell.Formula;
                                roomString = roomString.Replace("\"", "");

                                PageNameCell = visShapes.get_Cells(CUST_PROP_PREFIX + "DeviceShapePage");
                                PageNameCell.FormulaU = string.Format("\"{0}\"", IndexingPage.Name.ToString());
                                PageNameCell = visShapes.get_Cells(CUST_PROP_PREFIX + "DeviceShapeRoom");
                                PageNameCell.FormulaU = string.Format("\"{0}\"", roomString);
                            }
                            catch
                            {
                                if (f)
                                {
                                    newCustomProp.AddCustomProperty(visShapes, "DeviceShapeRoom", "Room", "Room Name", VisCellVals.visPropTypeString, "@", "", false, true, "");
                                    newCustomProp.AddCustomProperty(visShapes, "DeviceShapePage", "DeviceShapePage", "Page", VisCellVals.visPropTypeString, "@", "", false, true, "");
                                    pageSheet = document.Pages[PageToIndex.Index].PageSheet;
                                    tempPropertyCell = pageSheet.get_CellsU(CUST_PROP_PREFIX + "RoomName");
                                    roomString = tempPropertyCell.Formula;
                                    roomString = roomString.Replace("\"", "");
                                    PageNameCell = visShapes.get_Cells(CUST_PROP_PREFIX + "DeviceShapePage");
                                    PageNameCell.FormulaU = string.Format("\"{0}\"", IndexingPage.Name.ToString());
                                    PageNameCell = visShapes.get_Cells(CUST_PROP_PREFIX + "DeviceShapeRoom");
                                    PageNameCell.FormulaU = string.Format("\"{0}\"", roomString);
                                }
                            }
                        }
                    }
                }
                else
                {
                    break;
                }
            }

        }

        //SET DEVICE PAGE/DEVICE SCHEDULE
       
        //REVISION LOG RESET ENTRIES
        private void Button14_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Visio.Shape entryShape;
            Visio.Pages pages = document.Pages;
            Visio.Page page = pages[1];
            RevEntry revEntry = new RevEntry();
            int intRevX2;


            for (intRevX2 = page.Shapes.Count; intRevX2 >= 1; intRevX2--)
            {
                entryShape = page.Shapes[intRevX2];
                bool b = entryShape.Name.Contains("Revision");
                if (b)
                {
                    revEntry.DeleteCoverEntries(entryShape);

                }
            }
            revEntry.DropFullEntry();
        }

        //SET IP ADDRESSES
        private void Button15_Click(object sender, RibbonControlEventArgs e)
        {
            DeviceIPEntry newIPEntry = new DeviceIPEntry();
            newIPEntry.SetIPs();
        }

        //STRUCTURED WIRING SET
        private void Button16_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.Page currentPage = Globals.ThisAddIn.Application.ActivePage;
            string dropsStartingLetter = EditBox15.Text;
            string dropsStartingNumber = EditBox16.Text;
            double origX, origY;

            Dictionary<string, int> newDrops = new Dictionary<string, int>();
           
            newDrops.Add("[New] 1Port Net Drop", int.Parse(EditBox13.Text));
            newDrops.Add("[New] 2Port Net Drop", int.Parse(EditBox14.Text));
            newDrops.Add("[Existing] 1Port Net Drop", int.Parse(EditBox17.Text));
            newDrops.Add("[Existing] 2Port Net Drop", int.Parse(EditBox18.Text));
            if (CheckBox3.Checked)
            {
                newDrops.Add("[New] 2Port AP Drop", int.Parse(EditBox21.Text));
            }
            else
            {
                newDrops.Add("[New] 1Port AP Drop", int.Parse(EditBox21.Text));
            }

            StructuredWiringSymbols newStructWiring = new StructuredWiringSymbols();
            newStructWiring.newLetter = char.Parse(EditBox15.Text.Replace("\"", ""));
            newStructWiring.newNumber = int.Parse(EditBox16.Text);

            if(!CheckBox4.Checked)
                newStructWiring.CheckPage();

            
            origX = newStructWiring.x_coord;
            origY = newStructWiring.y_coord;

            foreach (KeyValuePair<string, int> kvp in newDrops)
            {
                for (int x = 1; x <= kvp.Value; x++)
                {
                    newStructWiring.SetNewDrops(kvp.Key);
                }
            }

            Visio.Shape backgroundRect = currentPage.DrawRectangle((origX+1.0), (origY+2.0), (newStructWiring.x_coord-1.0), newStructWiring.y_coordMax);
            backgroundRect.SendToBack();
            Visio.Characters textCharacters = backgroundRect.Characters;
            textCharacters.Text = "Drop Pool";
            textCharacters.set_CharProps(
                    (short)Microsoft.Office.Interop.Visio.VisCellIndices.
                    visCharacterStyle,
                    (short)Microsoft.Office.Interop.Visio.VisCellVals.
                    visSmallCaps);

            textCharacters.set_CharProps(
                    (short)Microsoft.Office.Interop.Visio.VisCellIndices.
                    visCharacterSize, 24);

            backgroundRect.get_CellsSRC(
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkVerticalAlign).FormulaU = "0";
        }

        //STRUCTURED WIRING ON SELECTED
        private void Button18_Click(object sender, RibbonControlEventArgs e)
        {
            string dropsStartingLetter = EditBox15.Text;
            string dropsStartingNumber = EditBox16.Text;
            StructuredWiringSymbols newStructWiring = new StructuredWiringSymbols();
            newStructWiring.newLetter = char.Parse(EditBox15.Text.Replace("\"", ""));
            newStructWiring.newNumber = int.Parse(EditBox16.Text);
            if(!CheckBox4.Checked)
                newStructWiring.CheckPage();
            newStructWiring.RunOnSelected();
        }

        //STRUCTURED WIRING SET
        private void Button20_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Page currentPage = Globals.ThisAddIn.Application.ActivePage;
            Visio.Documents visioDocuments = Globals.ThisAddIn.Application.Documents;
            Microsoft.Office.Interop.Visio.Selection shapesSelection;
            Visio.Document stencil;
            Visio.Master masterInStencil;
            Visio.Shape testingShape;
            Visio.Cell currentRevXCell, currentRevYCell;

            OpenDocumentSample openDoc;
            double entryX_location, y_Height;
            double x_coord, y_coord, x_coordMax, y_HeightMax;
            x_coord= .5;
            y_coord= 21.5;
            y_HeightMax = 1;
            int x = 0;
            List<string> missingDrops = new List<string>();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv"; // Filter for CSV files
            openFileDialog.Title = "Select a CSV File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string csvPath = openFileDialog.FileName;

                using (StreamReader reader = new StreamReader(csvPath))
                {
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] values = line.Split(','); // Assumes comma-separated values

                        // Process the values:
                        //foreach (string value in values)
                        //{
                            
                            //newDrops[x] = values;
                            //System.Diagnostics.Debug.WriteLine(newDrops[x].ToString());
                        string manu = values[0].ToString().ToUpper();
                        manu = manu + ".vssx";
                        try
                        {
                            stencil = visioDocuments[string.Format("\"{0}\"", manu)];
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {

                            openDoc = new OpenDocumentSample();
                            bool test = OpenDocumentSample.DemoDocumentOpen(visioApplication, manu);
                            System.Diagnostics.Debug.WriteLine(string.Format("Stencil Test \"{0}\" result", test));
                            // The stencil is not in the collection; open it as a 
                            // docked stencil.
                            try
                            {
                                stencil = visioDocuments.OpenEx(manu,
                                (short)Microsoft.Office.Interop.Visio.
                                VisOpenSaveArgs.visOpenDocked);
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Stencil\"{0}\"isn't available", manu));
                                stencil = null;
                                missingDrops.Add(values[1].ToString());
                                Console.ReadLine();
                                continue;
                            }

                        }
                        try
                        {
                            masterInStencil = stencil.Masters.get_ItemU(values[1]);
                            testingShape = masterInStencil.Shapes[1];
                            testingShape.Name = values[1].ToString();
                            shapesSelection = currentPage.CreateSelection(VisSelectionTypes.visSelTypeEmpty);
                            for (int y = 0; y < int.Parse(values[2]); y++)
                            {
                                testingShape = currentPage.Drop(masterInStencil, x_coord, y_coord);
                                testingShape.get_CellsU("LocPinX").FormulaU = string.Format("\"{0}\"", "Width*0");
                                testingShape.get_CellsU("LocPinY").FormulaU = string.Format("\"{0}\"", "Height*1");
                                currentRevXCell = testingShape.get_CellsU("Width");
                                entryX_location = currentRevXCell.get_Result("in");


                                currentRevYCell = testingShape.get_CellsU("Height");
                                y_Height = currentRevYCell.get_Result("in");
                                System.Diagnostics.Debug.WriteLine(values[1].ToString());

                                if (x_coord+entryX_location <= 29.5)
                                {
                                    //x_coordMax = x_coord;
                                    x_coord = x_coord + entryX_location + .25;
                                    if (y_Height > y_HeightMax)
                                    {
                                        y_HeightMax = y_Height;
                                    }
                                }
                                else
                                {
                                    x_coord = .5;
                                    y_coord = y_coord - y_HeightMax - .25;
                                    y_HeightMax = .5;
                                    testingShape.get_CellsU("PinX").FormulaU = string.Format("\"{0}\"", x_coord);
                                    testingShape.get_CellsU("PinY").FormulaU = string.Format("\"{0}\"", y_coord);
                                    x_coord = x_coord + entryX_location + .25;
                                    if (y_Height > y_HeightMax)
                                    {
                                        y_HeightMax = y_Height;
                                    }
                                }
                                shapesSelection.Select(testingShape, 0);
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            missingDrops.Add(values[1].ToString());
                        }
                        x++;
                    }
                }
            }

            Console.ReadLine();
            Globals.ThisAddIn.Application.ActiveDocument.Pages[1].Application.ActiveWindow.DeselectAll();
            string missingString = string.Join(", ", missingDrops);
            DialogResult dialogResult = MessageBox.Show(missingString, "Missing Drops");
        }

        private void SplitButton2_Click(object sender, RibbonControlEventArgs e)
        {
            this.splitButton2.Click -= new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.Button10_Click);
            this.splitButton2.Click -= new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.Button11_Click);
        }

        //Cable Label Numbering Textbox Change
        private void EditBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            int startingNum, incrementNum;

            if (this.editBox1.Text == "")
            {
                startingNum = 101;
            }
            else
            {
                startingNum = int.Parse(this.editBox1.Text);
            }
            if (this.checkBox2.Checked == false)
            {
                incrementNum = 1;
            }
            else
            {
                incrementNum = 100;
            }
            AutoNumbering("Cable_Num", startingNum, incrementNum, true);
        }
        //Account Details Textbox Change
        private void EditBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            DocumentDetailsSet("AccountDetails", editBox2.Text.ToString());
        }
        //Opportunity Textbox Change
        private void EditBox3_TextChanged(object sender, RibbonControlEventArgs e)
        {
            DocumentDetailsSet("OpportunityDetails", editBox3.Text.ToString());
        }
        //Job Number Textbox Change
        private void EditBox4_TextChanged(object sender, RibbonControlEventArgs e)
        {
            DocumentDetailsSet("JobNumSet", editBox4.Text.ToString());
            //JobNumSet(editBox4.Text);
        }
        //I/O Numbering Textbox Change
        private void EditBox6_TextChanged(object sender, RibbonControlEventArgs e)
        {
            int startingNum;
            if (this.editBox6.Text == "")
            {
                startingNum = 1;
            }
            else
            {
                startingNum = int.Parse(this.editBox6.Text);
            }
            AutoNumbering("IO_NUM", startingNum, 1, false);
        }
        //Device Number Textbox Change
        private void EditBox7_TextChanged(object sender, RibbonControlEventArgs e)
        {
            int startingNum;
            if (this.editBox7.Text == "")
            {
                startingNum = 1;
            }
            else
            {
                startingNum = int.Parse(this.editBox7.Text);
            }
            AutoNumbering("DeviceNum", startingNum, 1, true);
        }
        //Approved By Textbox Change
        private void EditBox8_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string strApprovedBy = "ApprovedBy";
            System.Diagnostics.Debug.WriteLine("editbox8 edit");
            TitlePageNamesSet(strApprovedBy, editBox8.Text.ToString());
        }
        //Designed By Textbox Change
        private void EditBox9_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string strDesignedBy = "DesignedBy";
            System.Diagnostics.Debug.WriteLine("editbox9 edit");
            TitlePageNamesSet(strDesignedBy, editBox9.Text.ToString());
        }
        //Drawn By Textbox Change
        private void EditBox10_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string strDrawnBy = "DrawnBy";
            System.Diagnostics.Debug.WriteLine("editbox10 edit");
            TitlePageNamesSet(strDrawnBy, editBox10.Text.ToString());
        }

        public void AutoNumbering(string rowNameU, int startingIncrementNum, int numberIncrement, bool fullNum)
        {
            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Window window = visioApplication.ActiveWindow;
            Visio.Selection theSelection;
            theSelection = window.Selection;

            const string CUST_PROP_PREFIX = "Prop.";
            try
            {
                foreach (Visio.Shape customPropertyShape in theSelection)
                {
                    string incrementedNumStr;

                    Visio.Cell customPropertyCell;

                    // Get the Cell object. Note the addition of "Prop." to the
                    // name given to the cell.
                    customPropertyCell = customPropertyShape.get_CellsU(CUST_PROP_PREFIX + rowNameU);
                    if (fullNum)
                    {
                        incrementedNumStr = startingIncrementNum.ToString("000");
                        customPropertyCell.FormulaU = "\"" + incrementedNumStr + "\"";
                    }
                    else
                    {
                        customPropertyCell.FormulaU =startingIncrementNum.ToString();
                    }
                    startingIncrementNum += numberIncrement;
                }
                return;
            }
            catch (System.Runtime.InteropServices.COMException)
            {

            }
        }
        public void DocumentDetailsSet(string strDetail, string strDetailText)
        {
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            //strDetail = 1;

            switch(strDetail)
            {
                case "AccountDetails":
                    document.Company = strDetailText;
                    break;
                case "OpportunityDetails":
                    document.Title = strDetailText;
                    break;
                case "JobNumSet":
                    document.Subject = strDetailText;
                    break;
            }
            
        }
        public void TitlePageNamesSet(string strNameBySetting, string strNameByText)
        {
            const string CUST_PROP_PREFIX = "Prop.";

            Visio.Page lastPage = Globals.ThisAddIn.Application.ActiveDocument.Pages[Globals.ThisAddIn.Application.ActiveDocument.Pages.Count];
            Visio.Cell customPropertyCell;

            try
            {

                System.Diagnostics.Debug.WriteLine("EnteredNameSetBy");
                foreach (Visio.Shape revShape in lastPage.Shapes)
                {
                    bool b = revShape.Name.Contains("TitleBlockNames");
                    if (b)
                    {
                        customPropertyCell = revShape.get_CellsU(CUST_PROP_PREFIX + strNameBySetting);
                        SetCellValueToString(customPropertyCell, strNameByText);
                        if(strNameBySetting == "DrawnBy")
                        {
                            Globals.ThisAddIn.Application.ActiveDocument.Creator = strNameByText;
                        }    
                        return;
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {

            }
        }
        public void TOCEntry(bool deleteOld)
        {
            Visio.Documents visioDocuments = Globals.ThisAddIn.Application.Documents;
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Visio.Document stencil;
            Visio.Master masterInStencil;
            Visio.Shape shapesInMaster;
            Visio.Shape testShape;
            Visio.Shape visShape;
            Visio.Shape pageSheet;
            Visio.Pages pages = document.Pages;
            Visio.Page page = pages[1];
            int intX2;
            double x_location, y_location, cellXDouble, cellYDouble;
            x_location = 0.00;
            y_location = 0.0;
            Visio.Cell currentXCell, currentYCell;
            const string CUST_PROP_PREFIX = "Prop.";


            for (intX2 = page.Shapes.Count; intX2 >= 1; intX2--)
            {
                visShape = page.Shapes[intX2];
                bool b = visShape.Name.Contains("TOC");
                if (b)
                {
                    currentXCell = visShape.get_CellsU("PinX");
                    cellXDouble = currentXCell.get_Result("in");

                    currentYCell = visShape.get_CellsU("PinY");
                    cellYDouble = currentYCell.get_Result("in");

                    if (cellXDouble >= x_location)
                    {
                        x_location = cellXDouble;
                    }
                    if (cellYDouble >= y_location)
                    {
                        y_location = cellYDouble;
                    }
                    visShape.Delete();
                }
            }

            try
            {
                stencil = visioDocuments["TitlePage Stencil.vssx"];
            }
            catch (System.Runtime.InteropServices.COMException)
            {

                // The stencil is not in the collection; open it as a 
                // docked stencil.
                stencil = visioDocuments.OpenEx("TitlePage Stencil.vssx",
                    (short)Microsoft.Office.Interop.Visio.
                    VisOpenSaveArgs.visOpenDocked);
            }

            // Get a master from the stencil by its universal name.
            foreach (Visio.Page eachPage in document.Pages)
            {
                if (eachPage.Background == page.Background)
                {
                    Visio.Page thisPage;
                    int intShapeCt, intX, x;
                    thisPage = eachPage;
                    intShapeCt = eachPage.Shapes.Count;
                    for (intX = 1; intX <= intShapeCt; intX++)
                    {
                        //if(checkShapes.Name.Contains("PageTitle"))
                        visShape = eachPage.Shapes[intX];
                        if (visShape.Name.Contains("PageTitle"))
                        {
                            Visio.Characters vsoCharacters, vsoCharacters1;
                            Visio.Cell tempPropertyCell;
                            masterInStencil = stencil.Masters.get_ItemU("TOC Entry");
                            testShape = masterInStencil.Shapes[1];
                            testShape.Name = "TOC Entry";
                            vsoCharacters1 = visShape.Characters;
                            testShape = page.Drop(testShape, x_location, y_location);
                            for (x = 1; x <= 4; x++)
                            {
                                shapesInMaster = testShape.Shapes[x];
                                vsoCharacters = shapesInMaster.Characters;
                                if (x == 1)
                                {
                                    vsoCharacters.Text = visShape.Text;
                                }
                                else if (x == 2)
                                {
                                    vsoCharacters.Text = eachPage.Name;
                                }
                                else if (x == 3)
                                {
                                    vsoCharacters.Text = eachPage.Index.ToString();
                                }
                                else if (x == 4)
                                {
                                    pageSheet = eachPage.PageSheet;
                                    try
                                    {
                                        tempPropertyCell = pageSheet.get_CellsU(CUST_PROP_PREFIX + "RoomName");
                                        string roomString = tempPropertyCell.Formula;
                                        roomString = roomString.Replace("\"", "");
                                        vsoCharacters.Text = roomString;
                                    }
                                    catch (System.Runtime.InteropServices.COMException)
                                    {
                                        vsoCharacters.Text = "---";
                                        
                                    }
                                }
                            }
                            testShape.get_CellsU("Hyperlink.CurPage.SubAddress").FormulaU = "=\"" + eachPage.Name + "\"";
                            testShape.get_CellsU("Hyperlink.CurPage.Description").FormulaU = "=\"" + eachPage.Name + "\"";
                            testShape.get_CellsU("TheText").FormulaU = "=0";
                            testShape.get_CellsU("User.visEquivTitle.Prompt").FormulaU = "=\"" + eachPage.Name + "\"";
                            y_location -= 0.5;
                        }
                    }
                }
            }
            page.Application.ActiveWindow.DeselectAll();
            stencil.Close();
            Globals.ThisAddIn.Application.BeforeDocumentClose += new EApplication_BeforeDocumentCloseEventHandler(NewDocument_Closed);
            return;
        }
        public void HostNaming(Visio.Shape equShape, string roomName)
        {
            Visio.Cell tempPropertyCell;
            const string CUST_PROP_PREFIX = "Prop.";
            string hostNameString;

            tempPropertyCell = equShape.get_CellsU(CUST_PROP_PREFIX + "DeviceName");
            hostNameString = tempPropertyCell.FormulaU;

            hostNameString = hostNameString + "-" + roomName;

            tempPropertyCell = equShape.get_CellsU(CUST_PROP_PREFIX + "DeviceNum");
            hostNameString = hostNameString + "-" + tempPropertyCell.Formula;

            hostNameString = hostNameString.Replace("\"", "");
            hostNameString = "\"" + hostNameString + "\"";

            tempPropertyCell = equShape.get_CellsU(CUST_PROP_PREFIX + "DeviceHostname");
            tempPropertyCell.FormulaForce = hostNameString;
        }
        public void SetCellValueToString(Microsoft.Office.Interop.Visio.Cell formulaCell, string newValue)
        {

            try
            {

                // Set the value for the cell.
                formulaCell.Formula = StringToFormulaForString(newValue);
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }
        }
        public static string StringToFormulaForString(string inputValue)
        {

            string result = "";
            string quote = "\"";
            string quoteQuote = "\"\"";

            try
            {

                result = inputValue != null ? inputValue : String.Empty;

                // Replace all (") with ("").
                result = result.Replace(quote, quoteQuote);

                // Add ("") around the whole string.
                result = quote + result + quote;
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }

            return result;
        }
        private void RevLevel_Label_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            document.Keywords = "1";
        }

        public class DeviceIPEntry
        {
            public Visio.Shape ipShape;
            public string acronymName, startIP, endIP;

            public static List<List<string>> ReadInIPs()
            {
                Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
                Visio.Document document = visioApplication.ActiveDocument;
                Visio.Pages visioPages = document.Pages;
                Visio.Shape shapesInEntryMaster;
                var acronymAR = new List<List<string>>();
                const string CUST_PROP_PREFIX = "Prop.";
                Visio.Cell tempBoolCell, tempValueCell;
                string acronymName, acronymStartingIP, acronymEndingIP;
                Visio.Characters vsoEntryCharacters;

                foreach (Visio.Shape CoverPageShapes in document.Pages[1].Shapes)
                {
                    bool b = CoverPageShapes.Name.Contains("AcronymEntry");
                    if (b)
                    {
                        shapesInEntryMaster = CoverPageShapes.Shapes[1];
                        vsoEntryCharacters = shapesInEntryMaster.Characters;
                        acronymName = vsoEntryCharacters.Text;
                        tempBoolCell = CoverPageShapes.get_CellsU(CUST_PROP_PREFIX + "HasIP");
                        if (System.Convert.ToBoolean(tempBoolCell.ResultIU))
                        {
                            tempValueCell = CoverPageShapes.get_CellsU(CUST_PROP_PREFIX + "FirstIPAddress");
                            acronymStartingIP = tempValueCell.FormulaU;
                            tempValueCell = CoverPageShapes.get_CellsU(CUST_PROP_PREFIX + "LastIPAddress");
                            acronymEndingIP = tempValueCell.FormulaU;
                            var acronymDict = new List<string>() { acronymName, acronymStartingIP, acronymEndingIP };
                            acronymAR.Add(acronymDict);
                        }
                    }
                }

                return acronymAR;
            }

            public void SetIPs()
            {
                Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
                Visio.Document document = visioApplication.ActiveDocument;
                Visio.Pages visioPages = document.Pages;
                Visio.Cell tempBoolCell;

                var readAcronymDict = new List<List<string>>();
                var readShapeDict = new Dictionary<Visio.Shape, string>();
                var myShapeCol = new List<Visio.Shape>();
                var ipShapeCol = new List<Visio.Shape>();
                var ipArrayList = new ArrayList();
                var ipNames = new List<DeviceIPEntry>();

                int x = 1;
                const string CUST_PROP_PREFIX = "Prop.";
                const string GATEWAY_STRING = "172.30.0.2";
                const string SCOUT_STRING = "172.30.1.250";

                readAcronymDict = ReadInIPs();

                while (document.Pages[x].Background == 0)
                {
                    foreach (Visio.Shape EquipmentShapes in document.Pages[x].Shapes)
                    {
                        bool b = EquipmentShapes.Name.Contains("Equipment Stencil");
                        if (b)
                        {
                            myShapeCol.Add(EquipmentShapes);
                        }
                    }
                    x++;
                }
                foreach (Visio.Shape testShape in myShapeCol)
                {
                    tempBoolCell = testShape.get_CellsU(CUST_PROP_PREFIX + "DeviceName");
                    foreach (var each in readAcronymDict)
                    {
                        if (each[0] == tempBoolCell.FormulaU.ToString().Replace("\"", ""))
                        {
                            ipNames.Add(new DeviceIPEntry() { ipShape = testShape, acronymName = each[0], startIP = each[1].Replace("\"", ""), endIP = each[2].Replace("\"", "") });
                        }
                    }
                }
                while (ipNames.Count > 0)
                {
                    List<DeviceIPEntry> ipResults = ipNames.FindAll(ipStart => ipStart.startIP == ipNames[0].startIP);
                    string lastOctetString;
                    int incOctet = 0;
                    foreach (var ip in ipResults)
                    {
                        //System.Diagnostics.Debug.WriteLine(ip.startIP + ", " + ip.endIP);
                        int lastOctet = int.Parse(ip.startIP.Substring(ip.startIP.LastIndexOf(".") + 1));
                        lastOctet = lastOctet + incOctet;
                        if (lastOctet <= int.Parse(ip.endIP.Substring(ip.endIP.LastIndexOf(".") + 1)))
                        {
                            lastOctetString = ip.startIP.Substring(0, ip.startIP.LastIndexOf(".") + 1) + lastOctet.ToString();
                            if (lastOctetString == GATEWAY_STRING || lastOctetString == SCOUT_STRING)
                            {
                                lastOctetString = ip.startIP.Substring(0, ip.startIP.LastIndexOf(".") + 1) + (lastOctet+=1).ToString();
                                ip.ipShape.get_CellsU(CUST_PROP_PREFIX + "DeviceIPAddress").Formula = "\"" + lastOctetString + "\"";
                                incOctet+=2;
                                ipNames.Remove(ip);
                            }
                            else
                            {
                                ip.ipShape.get_CellsU(CUST_PROP_PREFIX + "DeviceIPAddress").Formula = "\"" + lastOctetString + "\"";
                                incOctet++;
                                ipNames.Remove(ip);
                            }
                        }
                        else
                        {
                            ip.ipShape.get_CellsU(CUST_PROP_PREFIX + "DeviceIPAddress").Formula = "\"" + "0.0.0.0" + "\"";
                            ipNames.Remove(ip);
                        }
                    }
                }

            }


        }
        public class RevEntry
        {
            Visio.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            private Visio.Shape entryShape, shapesInEntryMaster;
            private Visio.Master masterInEntryStencil;
            private Visio.Characters vsoEntryCharacters;
            Visio.Cell currentRevXCell, currentRevYCell;

            private double entryX_location, entryY_location;
            public double RevNum;

            public string revType, revDescription;

            public bool boolMajorRev;

            public void SetShape(Visio.Shape settingShape)
            {
                entryShape = settingShape;

                currentRevXCell = entryShape.get_CellsU("PinX");
                entryX_location = currentRevXCell.get_Result("in");

                currentRevYCell = entryShape.get_CellsU("PinY");
                entryY_location = currentRevYCell.get_Result("in");

                shapesInEntryMaster = entryShape.Shapes[3];
                vsoEntryCharacters = shapesInEntryMaster.Characters;
                RevNum = double.Parse(vsoEntryCharacters.Text);
                entryY_location -= 0.5;
                if (boolMajorRev)
                {
                    RevNum = (int)RevNum + 1;
                }
                else
                {
                    RevNum += 0.1;
                }
                revDescription = "";

            }

            public void DropFullEntry()
            {
                Visio.Documents visioDocuments = Globals.ThisAddIn.Application.Documents;
                Visio.Document EntryStencil;

                try
                {
                    EntryStencil = visioDocuments["TitlePage Stencil.vssx"];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // The stencil is not in the collection; open it as a 
                    // docked stencil.
                    EntryStencil = visioDocuments.OpenEx("TitlePage Stencil.vssx",
                        (short)Microsoft.Office.Interop.Visio.
                        VisOpenSaveArgs.visOpenDocked);
                }

                masterInEntryStencil = EntryStencil.Masters.get_ItemU("Revision Entry");
                entryShape = Globals.ThisAddIn.Application.ActiveDocument.Pages[1].Drop(masterInEntryStencil, entryX_location, entryY_location);

                //REV reason 
                shapesInEntryMaster = entryShape.Shapes[1];
                vsoEntryCharacters = shapesInEntryMaster.Characters;
                vsoEntryCharacters.Text = revDescription;

                //Set test shape cell "Rev Date" to new rev date 
                shapesInEntryMaster = entryShape.Shapes[2];
                DateTime thisDay = DateTime.Today;
                var formatInfo = new CultureInfo("en-US").DateTimeFormat;
                formatInfo.DateSeparator = "/";
                shapesInEntryMaster.get_CellsU("Prop.RevDate").FormulaU = string.Format("\"{0}\"", thisDay.ToString("M/dd/yyyy", formatInfo));


                //Set test shape cell "RevNum" to new rev num 
                shapesInEntryMaster = entryShape.Shapes[3];
                vsoEntryCharacters = shapesInEntryMaster.Characters;
                vsoEntryCharacters.Text = RevNum.ToString();

                shapesInEntryMaster = entryShape.Shapes[4];
                vsoEntryCharacters = shapesInEntryMaster.Characters;
                vsoEntryCharacters.Text = revType;

                SetTitleEntry();
                Globals.ThisAddIn.Application.ActiveDocument.Pages[1].Application.ActiveWindow.DeselectAll();
                EntryStencil.Close();
            }

            public void SetTitleEntry()
            {
                //Set Title page current rev num
                Visio.Page lastPage = Globals.ThisAddIn.Application.ActiveDocument.Pages[Globals.ThisAddIn.Application.ActiveDocument.Pages.Count];

                foreach (Visio.Shape revShape in lastPage.Shapes)
                {
                    bool b = revShape.Name.Contains("revisionShape");
                    if (b)
                    {
                        shapesInEntryMaster = revShape.Shapes[1];
                        vsoEntryCharacters = shapesInEntryMaster.Characters;
                        vsoEntryCharacters.Text = RevNum.ToString();

                        //Set Title page current rev type
                        shapesInEntryMaster = revShape.Shapes[2];
                        vsoEntryCharacters = shapesInEntryMaster.Characters;
                        vsoEntryCharacters.Text = revType;
                        //

                        //Set Title page current rev date
                        shapesInEntryMaster = revShape.Shapes[4];
                        vsoEntryCharacters = shapesInEntryMaster.Characters;
                        DateTime thisDay = DateTime.Today; var formatInfo = new CultureInfo("en-US").DateTimeFormat;
                        formatInfo.DateSeparator = "/";
                        shapesInEntryMaster.get_CellsU("Prop.RevDate").FormulaU = string.Format("\"{0}\"", thisDay.ToString("M/dd/yyyy", formatInfo));
                        return;
                    }
                }
            }

            public void DeleteCoverEntries(Visio.Shape deletedRevShape)
            {
                entryShape = deletedRevShape;
                double tempXLocation, tempYLocation;

                currentRevXCell = entryShape.get_CellsU("PinX");
                tempXLocation = currentRevXCell.get_Result("in");

                currentRevYCell = entryShape.get_CellsU("PinY");
                tempYLocation = currentRevYCell.get_Result("in");

                if (entryX_location <= tempXLocation)
                {
                    entryX_location = tempXLocation;
                }
                if (entryY_location <= tempYLocation)
                {
                    entryY_location = tempYLocation;
                }

                entryShape.Delete();
                revType = "Build";
                RevNum = 0;
                revDescription = "Initial Build";
            }

        }
        public class StringValueInCell
        {

            /// <summary>This method is the class constructor.</summary>
            public StringValueInCell()
            {

                // No initialization is required.
            }

            /// <summary>This method sets the value of the specified Visio cell
            /// to the new string passed as a parameter.</summary>
            /// <param name="formulaCell">Cell in which the value is to be set
            /// </param>
            /// <param name="newValue">New string value that will be set</param>
            public void SetCellValueToString(
                Microsoft.Office.Interop.Visio.Cell formulaCell,
                string newValue)
            {

                try
                {

                    // Set the value for the cell.
                    formulaCell.FormulaU = StringToFormulaForString(newValue);
                }
                catch (Exception err)
                {
                    System.Diagnostics.Debug.WriteLine(err.Message);
                    throw;
                }
            }

            /// <summary>This method converts the input string to a Visio string by
            /// replacing each double quotation mark (") with a pair of double
            /// quotation marks ("") and then adding double quotation marks around
            /// the entire string.</summary>
            /// <param name="inputValue">Input string that will be converted
            /// to Visio String</param>
            /// <returns>Converted Visio string</returns>
            public static string StringToFormulaForString(string inputValue)
            {

                string result = "";
                string quote = "\"";
                string quoteQuote = "\"\"";

                try
                {

                    result = inputValue != null ? inputValue : String.Empty;

                    // Replace all (") with ("").
                    result = result.Replace(quote, quoteQuote);

                    // Add ("") around the whole string.
                    result = quote + result + quote;
                }
                catch (Exception err)
                {
                    System.Diagnostics.Debug.WriteLine(err.Message);
                    throw;
                }

                return result;
            }
        }
        public class AddingACustomProperty
        {

            /// <summary>This method is the class constructor.</summary>
            public AddingACustomProperty()
            {

                // No initialization is required.
            }

            /// <summary>This method creates a custom property for the shape that
            /// is passed in as a parameter.</summary>
            /// <param name="addedToShape">Shape to which the custom property is 
            /// to be added</param>
            /// <param name="localRowName">Local name for the row. This name will
            /// appear in custom properties dialog for users running in developer
            /// mode.</param>
            /// <param name="rowNameU">Universal name of the custom property to be
            /// created</param>
            /// <param name="labelName">Label of the custom property</param>
            /// <param name="propType">Type of the value of the custom property.
            /// Not all VisCellVals constants are valid for this parameter. Only
            /// constants that start with visPropType make sense in this context.
            /// </param>
            /// <param name="format">Format of the custom property</param>
            /// <param name="prompt">Prompt for the custom property</param>
            /// <param name="askOnDrop">Value of the "Ask On Drop" check box of the
            /// custom property. Only seen in developer mode</param>
            /// <param name="hidden">Value of the "Hidden" check box of the custom
            /// property. Only seen in developer mode</param>
            /// <param name="sortKey">Value of the Sort key of the custom property.
            /// Only seen in developer mode</param>
            /// <returns>True if successful; false otherwise</returns>
            public bool AddCustomProperty(
                Microsoft.Office.Interop.Visio.Shape addedToShape,
                string localRowName,
                string rowNameU,
                string labelName,
                Microsoft.Office.Interop.Visio.VisCellVals propType,
                string format,
                string prompt,
                bool askOnDrop,
                bool hidden,
                string sortKey)
            {

                const string CUST_PROP_PREFIX = "Prop.";

                Microsoft.Office.Interop.Visio.Cell shapeCell;
                short rowIndex;
                StringValueInCell formatHelper;
                bool returnValue = false;


                if (addedToShape == null)
                {
                    return false;
                }

                try
                {

                    formatHelper = new StringValueInCell();

                    // Add a named custom property row. In addition to adding a row
                    // with the local name, specified via localRowName parameter,
                    // this call will usually set the universal name of the new row
                    // to localRowName as well. However, the universal row name 
                    // will not be set if this shape already has a custom property 
                    // row that has the universal name equal to localRowName.
                    rowIndex = addedToShape.AddNamedRow(
                        (short)(Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp),
                        localRowName, (short)(Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowProp));

                    // The columns of the properties to set are fixed and can be
                    // accessed directly using the CellsSRC method and column index.

                    // Get the Cell object for each one of the items in the
                    // custom property and set its value using the FormulaU property
                    // of the Cell object.

                    // Column 1 : Prompt
                    shapeCell = addedToShape.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp, rowIndex,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visCustPropsPrompt);

                    formatHelper.SetCellValueToString(shapeCell,
                        prompt);

                    // Any cell in the row can be used to set the universal
                    // row name. Only set the name if rowNameU parameter differs
                    // from the local name and is not blank.
                    if (rowNameU != null)
                    {
                        if ((localRowName != rowNameU) && (rowNameU.Length > 0))
                        {
                            shapeCell.RowNameU = rowNameU;
                        }
                    }

                    // Column 2 : Label
                    shapeCell = addedToShape.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp, rowIndex,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visCustPropsLabel);

                    formatHelper.SetCellValueToString(shapeCell,
                        labelName);

                    // Column 3 : Format
                    shapeCell = addedToShape.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp, rowIndex,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visCustPropsFormat);

                    formatHelper.SetCellValueToString(shapeCell,
                        format);

                    // Column 4 : Sort Key
                    shapeCell = addedToShape.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp, rowIndex,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visCustPropsSortKey);

                    formatHelper.SetCellValueToString(shapeCell,
                        sortKey);

                    // Column 5 : Type
                    shapeCell = addedToShape.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp, rowIndex,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visCustPropsType);

                    formatHelper.SetCellValueToString(shapeCell,
                        ((short)propType).ToString(
                            System.Globalization.CultureInfo.InvariantCulture));

                    // Column 6 : Hidden (This corresponds to the invisible cell in
                    // the Shapesheet)
                    shapeCell = addedToShape.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp, rowIndex,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visCustPropsInvis);

                    formatHelper.SetCellValueToString(shapeCell,
                        hidden.ToString(
                            System.Globalization.CultureInfo.InvariantCulture));

                    // Column 7 : Ask on drop
                    shapeCell = addedToShape.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionProp, rowIndex,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visCustPropsAsk);

                    formatHelper.SetCellValueToString(shapeCell,
                        askOnDrop.ToString(
                            System.Globalization.CultureInfo.InvariantCulture));

                    // Set the custom property for the shape using FormulaU 
                    // property of the cell.
                    shapeCell = addedToShape.get_CellsU(CUST_PROP_PREFIX
                        + rowNameU);
                    formatHelper.SetCellValueToString(shapeCell,
                        rowNameU);

                    returnValue = true;
                }
                catch (Exception err)
                {
                    System.Diagnostics.Debug.WriteLine(err.Message);
                    throw;
                }

                return returnValue;
            }
        }

        public class StructuredWiringSymbols
        {
            public char newLetter = 'A';
            public int newNumber = 0;
            public double x_coord = -2.0;
            public double y_coord = 20.0;
            public double y_coordMax;

            public void CheckPage()
            {
                const string CUST_PROP_PREFIX = "Prop.";
                Visio.Page currentPage = Globals.ThisAddIn.Application.ActivePage;
                Visio.Cell firstPortLtrCheck, secondPortLtrCheck, firstPortNumCheck, secondPortNumCheck;

                foreach (Visio.Shape checkedShape in currentPage.Shapes)
                {
                    bool b = checkedShape.Name.Contains("Net Drop");
                    if (b)
                    {
                        try
                        {
                            secondPortLtrCheck = checkedShape.get_CellsU(CUST_PROP_PREFIX + "2ndPortLetter");
                            secondPortNumCheck = checkedShape.get_CellsU(CUST_PROP_PREFIX + "2ndPortNumber");
                            if (char.Parse(secondPortLtrCheck.Formula.Replace("\"", "")) > newLetter)
                            {
                                newLetter = char.Parse(secondPortLtrCheck.Formula.Replace("\"", ""));
                                newNumber = int.Parse(secondPortNumCheck.Formula.Replace("\"", "") + 1);
                            }
                        }
                        catch
                        {
                            firstPortLtrCheck = checkedShape.get_CellsU(CUST_PROP_PREFIX + "1stPortLetter");
                            firstPortNumCheck = checkedShape.get_CellsU(CUST_PROP_PREFIX + "1stPortNumber");
                            if (char.Parse(firstPortLtrCheck.Formula.Replace("\"", "")) > newLetter)
                            {
                                newLetter = char.Parse(firstPortLtrCheck.Formula.Replace("\"", ""));
                                newNumber = int.Parse(firstPortNumCheck.Formula.Replace("\"", ""))+1;
                            }
                        }
                    }
                }
                return;
            }

            public void RunOnSelected()
            {
                Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
                Visio.Window window = visioApplication.ActiveWindow;
                Visio.Selection theSelection;
                theSelection = window.Selection;
                const string CUST_PROP_PREFIX = "Prop.";

                try
                {
                    foreach (Visio.Shape SelectedShape in theSelection)
                    {
                        bool b = SelectedShape.Name.Contains("Drop");
                        if (b)
                        {
                            Visio.Cell portCellLetter, portCellNumber;
                            portCellLetter = SelectedShape.get_CellsU(CUST_PROP_PREFIX + "1stPortLetter");
                            portCellNumber = SelectedShape.get_CellsU(CUST_PROP_PREFIX + "1stPortNumber");
                            portCellLetter.FormulaU = "\"" + newLetter.ToString() + "\"";
                            portCellNumber.FormulaU = "\"" + newNumber.ToString("00") + "\"";
                            IncrementValues(1.5);
                            try
                            {
                                portCellLetter = SelectedShape.get_CellsU(CUST_PROP_PREFIX + "2ndPortLetter");
                                portCellNumber = SelectedShape.get_CellsU(CUST_PROP_PREFIX + "2ndPortNumber");
                                portCellLetter.FormulaU = "\"" + newLetter.ToString() + "\"";
                                portCellNumber.FormulaU = "\"" + newNumber.ToString("00") + "\"";
                                IncrementValues(0.5);
                            }
                            catch
                            {
                            }
                        }
                    }
                    return;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                }
            }

            public void SetNewDrops(string dropType)
            {
                const string CUST_PROP_PREFIX = "Prop.";
                Visio.Documents visioDocuments = Globals.ThisAddIn.Application.Documents;
                Visio.Page currentPage = Globals.ThisAddIn.Application.ActivePage;
                Visio.Document stencil;
                Visio.Master masterInStencil;
                Visio.Cell portCellLetter, portCellNumber;
                Visio.Shape testingShape;

                try
                {
                    stencil = visioDocuments["Structured Wiring.vssm"];
                }
                catch (System.Runtime.InteropServices.COMException)
                {

                    // The stencil is not in the collection; open it as a 
                    // docked stencil.
                    stencil = visioDocuments.OpenEx("Structured Wiring.vssm",
                        (short)Microsoft.Office.Interop.Visio.
                        VisOpenSaveArgs.visOpenDocked);
                }
                masterInStencil = stencil.Masters.get_ItemU(dropType);
                testingShape = masterInStencil.Shapes[1];
                testingShape.Name = dropType;
                testingShape = currentPage.Drop(testingShape, x_coord, y_coord);
                portCellLetter = testingShape.get_CellsU(CUST_PROP_PREFIX + "1stPortLetter");
                portCellNumber = testingShape.get_CellsU(CUST_PROP_PREFIX + "1stPortNumber");
                portCellLetter.FormulaU = "\"" + newLetter.ToString() + "\"";
                portCellNumber.FormulaU = "\"" + newNumber.ToString("00") + "\"";
                IncrementValues(1.5);
                try
                {
                    portCellLetter = testingShape.get_CellsU(CUST_PROP_PREFIX + "2ndPortLetter");
                    portCellNumber = testingShape.get_CellsU(CUST_PROP_PREFIX + "2ndPortNumber");
                    portCellLetter.FormulaU = "\"" + newLetter.ToString() + "\"";
                    portCellNumber.FormulaU = "\"" + newNumber.ToString("00") + "\"";
                    IncrementValues(0.5);
                }
                catch
                {
                }
            }
            private void IncrementValues(double dropDist)
            {
                if (newNumber < 24)
                {
                    newNumber++;
                }
                else
                {
                    newLetter++;
                    if (newLetter=='I' || newLetter == 'O')
                    {
                        newLetter++;
                        newNumber = 1;
                    }
                    else
                    {
                        newNumber = 1;
                    }
                }

                y_coord -= dropDist;
                
                if (y_coord <= 0)
                {
                    y_coordMax = y_coord;
                    x_coord -= 1.5;
                    y_coord = 20;
                }
                else
                {
                    if (y_coordMax >= y_coord)
                        y_coordMax = y_coord;
                }
                
            }
        }
       

        private void EditBox11_TextChanged_1(object sender, RibbonControlEventArgs e)
        {
        }

        private Visio.Shape GetShape(string pageUID, string shapeUID)
        {
            //Visio.Page pag;
            Visio.Shape shp;

            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Document document = visioApplication.ActiveDocument;
            foreach (Visio.Page pag in document.Pages)
            {
                if (pag.PageSheet.get_UniqueID((short)VisUniqueIDArgs.visGetOrMakeGUID) == pageUID)
                {
                    shp = pag.Shapes[shapeUID];
                    return shp;
                }
            }
            return null;
        }

        private Visio.Master GetMasterU(string masterName)
        {
            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Document document = visioApplication.ActiveDocument;

            if (masterName.Length == 0)
            {
                return null;
            }

            foreach(Visio.Master mst in document.Masters)
            {
                if(mst.NameU.ToUpper() == masterName.ToUpper()) 
                {
                    return mst;
                }
            }
            return null;
        }

        private void Button19_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.Application visioApplication = Globals.ThisAddIn.Application.Application;
            Visio.Document document = visioApplication.ActiveDocument;
            const string MSTR = "Pulled Cable V3 Broadcast";

            Visio.Page pag;
            Visio.Page trgtPage;
            Visio.Shape trgtShape;
            Visio.Characters chars;
            Visio.Master mst;

            Visio.Pages visioPages = document.Pages;
            Visio.Page page = visioApplication.ActivePage;

            string trgtPageUID, trgtShapeUID, sendingString; // receivingIO, sendingIO, sendingTo, sendingFrom;

            DialogResult dialogResult = MessageBox.Show("Are you sure?", "Set Flag Assignment", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                mst = GetMasterU(MSTR);
                pag = visioApplication.ActivePage;
                foreach (Visio.Shape visShapes in pag.Shapes)
                {
                    //Visio.Selection sel = pag.CreateSelection(VisSelectionTypes.visSelTypeByMaster, 0, mst);
                    bool b = visShapes.Name.Contains("Pulled Cable");
                    if (b)
                    {
                        try
                        {
                            if (visShapes.get_CellExistsU("User.OPCDPageID", (short)Visio.VisExistsFlags.visExistsAnywhere) != 0)
                            {
                                trgtPageUID = visShapes.get_CellsU("User.OPCDPageID").get_ResultStr("");
                                trgtShapeUID = visShapes.get_CellsU("User.OPCDShapeID").get_ResultStr("");
                                if ((trgtPageUID.Length == 38) && (trgtShapeUID.Length == 38))
                                {
                                    trgtShape = GetShape(trgtPageUID, trgtShapeUID);
                                    if (trgtShape != null)
                                    {
                                        trgtPage = trgtShape.ContainingPage;
                                        if (visShapes.get_CellExistsU("User.visEquivTitle", (short)Visio.VisExistsFlags.visExistsAnywhere) != 0)
                                        {
                                            visShapes.get_CellsU("Hyperlink.OffPageConnector.SubAddress").FormulaU = "=\"" + trgtPage.Name + "\"";
                                            visShapes.get_CellsU("Hyperlink.OffPageConnector.Description").FormulaU = "=User.visEquivTitle";
                                            visShapes.get_CellsU("TheText").FormulaU = "=0";
                                            visShapes.get_CellsU("User.visEquivTitle.Prompt").FormulaU = "=\"" + trgtPage.Name + "\"";

                                            if (visShapes.get_CellExistsU("User.visEquivIO_PORT", (short)Visio.VisExistsFlags.visExistsAnywhere) != 0)
                                            {
                                                //Set Far End IO to Page Flag
                                                visShapes.get_CellsU("User.visEquivIO_PORT").FormulaU = "=Pages[" + trgtPage.Name + "]!Sheet." + trgtShape.ID + "!Prop.IO";
                                                visShapes.get_CellsU("User.visEquivIO_NUM").FormulaU = "=Pages[" + trgtPage.Name + "]!Sheet." + trgtShape.ID + "!Prop.IO_NUM";

                                                //Set Page Flag to Far End
                                                sendingString = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID + "!Prop.Cable_Num";
                                                trgtShape.get_CellsU("User.visEquivIO_PORT").Formula = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID.ToString() + "!Prop.IO";
                                                trgtShape.get_CellsU("User.visEquivIO_NUM").Formula = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID.ToString() + "!Prop.IO_NUM";
                                            }

                                            trgtShape.get_CellsU("Prop.Cable_Label").Formula = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID + "!Prop.Cable_Label";
                                            trgtShape.get_CellsU("Prop.Cable_Num").Formula = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID.ToString() + "!Prop.Cable_Num";
                                            trgtShape.get_CellsU("Prop.Cable_Type").Formula = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID.ToString() + "!Prop.Cable_Type";
                                            trgtShape.get_CellsU("Prop.Cable_To").Formula = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID + "!Prop.Cable_To";
                                            trgtShape.get_CellsU("Prop.Cable_From").Formula = "=Pages[" + pag.Name + "]!Sheet." + visShapes.ID.ToString() + "!Prop.Cable_From";
                                            trgtShape.get_CellsU("Prop.Cable_HashTag").FormulaU = "=\"" + "XX" + "\"";

                                            chars = visShapes.Characters;
                                            chars.Delete();
                                            chars.Begin = 0;
                                            chars.End = 1;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception err)
                        {
                            System.Diagnostics.Debug.WriteLine(err.Message);
                        }
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                
            }

        }

    }
    public class OpenDocumentSample
    {

        /// <summary>This constructor is intentionally left blank.</summary>
        public OpenDocumentSample()
        {

            // no initialization required.
        }

        /// <summary>This method opens three files to demonstrate some of the
        /// options available to open Visio files.</summary>
        /// <param name="applicationObj">Visio instance to open the
        /// files in</param>
        /// <param name="stencilName">Name of a Visio stencil file (including
        /// the full path) that will be opened docked and in read-only mode
        /// </param>
        /// <param name="documentName">Name of a Visio document file
        /// (including the full path) of which a copy will be opened</param>
        /// <param name="hiddenDocumentName">Name of a Visio document
        /// file (including the full path) which will be opened as a hidden
        /// drawing, that will have macros disabled, and will not appear on
        /// the recent files list</param>
        /// <returns>True if all three files are opened successfully</returns>
        public static bool DemoDocumentOpen(
            Microsoft.Office.Interop.Visio.Application applicationObj,
            string stencilName)
        {

            bool documentsOpened = false;

            if (applicationObj == null)
            {
                return false;
            }

            try
            {

                // Only open the files if all of the files exist.
                if (System.IO.File.Exists(stencilName))
                {

                    // Open the files.  The OpenEx method will raise an
                    // error if the document cannot be opened.  The
                    // flags set in the second argument determine the
                    // properties applied when opening the document.
                    applicationObj.Documents.OpenEx(stencilName,
                        ((short)Microsoft.Office.Interop.Visio.
                            VisOpenSaveArgs.visOpenDocked +
                        (short)Microsoft.Office.Interop.Visio.
                            VisOpenSaveArgs.visOpenRO));



                    documentsOpened = true;
                }
            }

            catch (Exception error)
            {
                System.Diagnostics.Debug.WriteLine(error.Message);
                throw;
            }

            return documentsOpened;
        }
    }


}
