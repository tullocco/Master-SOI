using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing.Printing;

namespace Master_SOI
{
    public partial class StriteRevisionTab
    {
        string Address = @"P:\ISO Documents\ISO MASTER SOI(AUTOMATED)\SOI-";
        string MasterAddress = @"P:\ISO Documents\ISO MASTER SOI(AUTOMATED)\Blank SOI-Master.docx";                                         // Master address also in ThisDocument.cs -> Doc.PopSOISelect()
        //string MasterSaveAddress = @"H:\ISO Documents\ISO MASTER SOI(AUTOMATED)\SOI-Master.docx";

        List<string> revMemory = new List<string>(); 

        TrackingSheetEditorPane TrkShtPane = new TrackingSheetEditorPane();

        private void StriteRevisionTab_Load(object sender, RibbonUIEventArgs e)
        {
            Doc.PopSOISelect(); 

            Globals.ThisDocument.ActionsPane.Controls.Add(TrkShtPane);
            TrkShtPane.Hide();
            Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = false;

            NewSOI.Click += NewSOI_Click;
            SOISelect.SelectionChanged += SOISelect_SelectionChanged;            
            EditTrackSht.Click += EditTrackSht_Click;
            NewRev.Click += NewRev_Click;
            RevSelect.SelectionChanged += RevSelect_SelectionChanged;
            AcceptRev.Click += AcceptRev_Click;
            RejectRev.Click += RejectRev_Click;
            Review.Click += Review_Click;
            Submit.Click += Submit_Click;
            PrintRev.Click += PrintRev_Click;

        }

        private void PrintRev_Click(object sender, RibbonControlEventArgs e)
        {            
            if (RevSelect.Items.Count > 0)
            {
                string[] LTR = RevSelect.SelectedItem.Label.Split(' ');
                if (LTR[0] == "TBA")
                {
                    Globals.ThisDocument.TrackRevisions = false;
                    Globals.ThisDocument.FooterLTR.Text = LTR[1];
                    Globals.ThisDocument.FooterLTR2.Text = LTR[1];
                    Globals.ThisDocument.LTR.Text = LTR[1];

                    Globals.ThisDocument.Application.ActivePrinter = "HP Officejet Pro 8600 (Network)";

                    object what = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage;                                      //Go to First Pg
                    object which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst;
                    object count = 1;
                    object missing = System.Reflection.Missing.Value;

                    Globals.ThisDocument.Application.Selection.GoTo(ref what, ref which, ref count, ref missing);

                    object copies = "1";                            
                    object pages = "1";
                    object range = Word.WdPrintOutRange.wdPrintCurrentPage;
                    object items = Word.WdPrintOutItem.wdPrintDocumentContent;
                    object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                    object oTrue = true;
                    object oFalse = false;

                    Globals.ThisDocument.PrintOut(ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,  //Print 1st pg
                        ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                        ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);

                    Globals.ThisDocument.Application.ActivePrinter = @"\\DC03.STRITELTD.NET\Canon iR-ADV 4051 UFR II";                    

                    pages = "";
                    range = Word.WdPrintOutRange.wdPrintAllDocument;

                    Globals.ThisDocument.PrintOut(ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,  //Print whole doc
                        ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                        ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);

                    Globals.ThisDocument.FooterLTR.Text = RevSelect.SelectedItem.Label;
                    Globals.ThisDocument.FooterLTR2.Text = RevSelect.SelectedItem.Label;
                    Globals.ThisDocument.LTR.Text = RevSelect.SelectedItem.Label;
                    Globals.ThisDocument.TrackRevisions = true;
                }
            }   
        }

        private void Submit_Click(object sender, RibbonControlEventArgs e)
        {
            if(NewRev.Checked == true)
            {
                DialogResult result = MessageBox.Show("Are You Sure You Want to Submit These Changes?", "Submit Changes?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    NewRev.Checked = false;
                    Globals.ThisDocument.TrackRevisions = false;

                    object start = Globals.ThisDocument.Content.End - 1;
                    object end = Globals.ThisDocument.Content.End;
                    Word.Range AddSpace = Globals.ThisDocument.Range(ref start, ref end);                                                       // Adds a space to the end of the document incase there are no revision bars to delete
                    AddSpace.Text = " ";

                    string revMem = Globals.ThisDocument.FooterSOINum.Text + " " + Globals.ThisDocument.FooterLTR.Text;
                    revMemory.Add(revMem);

                    string[] LTR = RevSelect.SelectedItem.Label.Split(' ');                                                                     // Gets Rid of TBA

                    RevSelect.SelectedItem.Label = LTR[1];
                    Globals.ThisDocument.LTR.Text = LTR[1];

                    Doc.Protect();

                    SaveFile();
                }   
            }
        }

        private void Review_Click(object sender, RibbonControlEventArgs e)
        {
            if (NewRev.Checked == true)
            {
                Globals.ThisDocument.TrackRevisions = false;

                DateTime AuthDate = DateTime.Now;

                Globals.ThisDocument.DateRevised.Text = AuthDate.ToString("MM/dd/yy");
                Globals.ThisDocument.FooterDateRevised.Text = AuthDate.ToString("MM/dd/yy");
                Globals.ThisDocument.FooterDateRevised2.Text = AuthDate.ToString("MM/dd/yy");
                Globals.ThisDocument.Auth.Text = "N.H";

                Globals.ThisDocument.TrackRevisions = true;                
                SaveFile();
            }
            
        }

        private void RejectRev_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisDocument.RejectAllRevisions();

            Word.Revision ThisRevision = Globals.ThisDocument.ActiveWindow.Selection.NextRevision();
                        
            //Word.Range RevRange = ThisRevision.Range;

            ThisRevision.Reject();

        }

        private void AcceptRev_Click(object sender, RibbonControlEventArgs e)
        {   
            if(NewRev.Checked == true)
            {  
                int index = 0;
               
                string num = " ";
                string clauseNum = " ";

                List<string> checkNum = new List<string>();

                Word.Revision ThisRevision = Globals.ThisDocument.ActiveWindow.Selection.NextRevision();
                Word.Range StartYRng = Globals.ThisDocument.ActiveWindow.Selection.Range;                                                  // Gets Start Range for Revision Bar 

                Globals.ThisDocument.ActiveWindow.Selection.MoveDown();
                Word.Range EndYRng = Globals.ThisDocument.ActiveWindow.Selection.Range;                     // Gets End Range for Revision Bar

                Word.Paragraphs para = Globals.ThisDocument.Paragraphs;                                     

                int j = 1;

                while (j <= para.Count)                                                                     // Searches thru paragraphs to see if revision is in them
                {
                    Word.Range parRange = Globals.ThisDocument.Paragraphs[j].Range;
                    StartYRng.Select();
                    bool InRange = Globals.ThisDocument.ActiveWindow.Selection.InRange(parRange);
                    parRange.Select();
                    bool InRange2 = Globals.ThisDocument.ActiveWindow.Selection.InRange(StartYRng);
                    if (InRange == true || InRange2 == true)
                    {
                        char[] delimiterChars = { ' ', '\t','\r','\f'};
                        string[] parText = parRange.Text.Split(delimiterChars);

                        string ClauseText = parText[0];

                        char num1 = ClauseText[0];
                        char pt = ClauseText[1];
                        char num2 = ClauseText[2];

                        bool isNum1 = System.Char.IsDigit(num1);
                        bool isNum2 = System.Char.IsDigit(num2);
                        
                        while (isNum1 == false || isNum2 == false)                                          // If clause number isnt at the front of the paragraph, move paragraphs until it is
                        {
                            j--;
                            if(j == 0)
                            {
                                while(isNum1 == false || isNum2 == false)
                                {
                                    j++;
                                    parRange = Globals.ThisDocument.Paragraphs[j].Range;                                    
                                    parText = parRange.Text.Split(delimiterChars);

                                    ClauseText = parText[0];
                                    num1 = ClauseText[0];
                                    pt = ClauseText[1];
                                    num2 = ClauseText[2];

                                    isNum1 = System.Char.IsDigit(num1);
                                    isNum2 = System.Char.IsDigit(num2);

                                }
                            }
                            else
                            {
                                parRange = Globals.ThisDocument.Paragraphs[j].Range;                                
                                parText = parRange.Text.Split(delimiterChars);

                                ClauseText = parText[0];
                                num1 = ClauseText[0];
                                pt = ClauseText[1];
                                num2 = ClauseText[2];

                                isNum1 = System.Char.IsDigit(num1);
                                isNum2 = System.Char.IsDigit(num2);
                            }                            
                        }

                        num = $"{num1}" + $"{pt}" + $"{num2}";
                        checkNum.Add(num);

                        int listLength = checkNum.Count;

                        if (listLength > 1)
                        {
                            string prevNum = checkNum[index - 1];
                            if (num != prevNum)
                                clauseNum = clauseNum + num + ", ";                            
                        }
                        else
                        {
                            clauseNum = clauseNum + num + ", ";
                        }
                        ThisRevision.Accept();
                        j = para.Count + 1;
                        index++;
                    }
                    else
                    {
                        j++;
                    }
                }

                float StartX = 40;
                float StartY = StartYRng.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
                float EndX = 40;
                float EndY = EndYRng.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];

                Globals.ThisDocument.TrackRevisions = false;
                Globals.ThisDocument.Shapes.AddLine(StartX, StartY, EndX, EndY);
                Globals.ThisDocument.Description.Text = Globals.ThisDocument.Description.Text + "Revised Sections" + clauseNum;
                Globals.ThisDocument.TrackRevisions = true;                   
            }            
        }

        private void RevSelect_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            string[] LTR = RevSelect.SelectedItem.Label.Split(' ');
            if (LTR[0] == "TBA")                                                                // Unprotect Doc or turn off rev tracking
            {
                DialogResult result = MessageBox.Show("Save Changes to SOI?", "Save?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                    SaveFile();

                Globals.ThisDocument.TrackRevisions = false;
            }

            if (NewRev.Checked == false)
                Doc.UnProtect();
            else
                SaveFile();
            Globals.ThisDocument.TrackRevisions = false;

            Doc.UnProtect();
            string[] LoadTitle = SOISelect.SelectedItem.Label.Split(' ');
            Word.Range Document = Globals.ThisDocument.Content;
            int i = 0;
            string Title = "";
            foreach (var text in LoadTitle)
            {
                if (i != 0)
                {
                    Title = Title + " " + text;
                }
                else
                    i++;
            }

            string docName = Address + Globals.ThisDocument.FooterSOINum.Text + " " + Globals.ThisDocument.Title.Text + @"\SOI-" + LoadTitle[0] + " " + "(" + RevSelect.SelectedItem.Label + ")" + Title + ".docx";
            Document.ImportFragment(docName);
            Globals.ThisDocument.FooterDateRevised.Text = Globals.ThisDocument.DateRevised.Text;
            Globals.ThisDocument.FooterDateRevised2.Text = Globals.ThisDocument.DateRevised.Text;
            Globals.ThisDocument.FooterLTR.Text = Globals.ThisDocument.LTR.Text;
            Globals.ThisDocument.FooterLTR2.Text = Globals.ThisDocument.LTR.Text;

            if (LTR[0] == "TBA")
                Globals.ThisDocument.TrackRevisions = true;
            else
                Doc.Protect();
        }

        private void NewRev_Click(object sender, RibbonControlEventArgs e)
        {
            
            object start = Globals.ThisDocument.Content.End - 1;
            object end = Globals.ThisDocument.Content.End;
            string LTR;

            if (NewRev.Checked == true)
            {
                Doc.UnProtect();                
                Globals.ThisDocument.Range(ref start, ref end).Select();
                Globals.ThisDocument.Shapes.SelectAll();                                                                                // Selects Revision Bars
                Globals.ThisDocument.ActiveWindow.Selection.Delete();                                                                   // Deletes Rev Bars
                Globals.ThisDocument.TrackRevisions = false;
                Globals.ThisDocument.ShowRevisions = true;

                Globals.ThisDocument.Description.Text = " ";

                LoadFile();

                LTR = Globals.ThisDocument.LTR.Text;
                if(LTR == "")
                {
                    LTR = "TBA A";
                }
                else
                {
                    int i = 0;
                    char LTR1 = ' ';
                    char LTR2 = ' ';
                    bool AddTwo = false;

                    foreach (var Char in LTR)
                    {
                        if (i == 0)
                        {
                            if (Char == 'Z')
                            {
                                LTR1 = 'A';
                                AddTwo = true;
                            }
                            else
                            {
                                LTR1 = Char;
                                LTR1++;
                                AddTwo = false;
                            }
                            LTR = "TBA " + $"{LTR1}";
                        }
                        else
                        {
                            if (AddTwo == true)
                            {
                                LTR2 = Char;
                            }
                            else
                            {
                                LTR2 = Char;
                                LTR2++;
                            }
                            LTR = "TBA " + $"{LTR1}" + $"{LTR2}";
                        }
                    }
                }                

                Globals.ThisDocument.LTR.Text = LTR;
                Globals.ThisDocument.FooterLTR.Text = LTR;
                Globals.ThisDocument.FooterLTR2.Text = LTR;
                Globals.ThisDocument.DateRevised.Text = "";
                Globals.ThisDocument.FooterDateRevised.Text = "";
                Globals.ThisDocument.FooterDateRevised2.Text = "";
                Globals.ThisDocument.Auth.Text = "";

                RibbonDropDownItem RevObj = Factory.CreateRibbonDropDownItem();
                RevObj.Label = LTR;
                
                RevSelect.Items.Add(RevObj);

                int j = RevSelect.Items.Count;
                RevSelect.SelectedItem = RevObj;

                Globals.ThisDocument.TrackRevisions = true;
            }
            else
            {
                NewRev.Checked = true;
            }
        }
    
        private void EditTrackSht_Click(object sender, RibbonControlEventArgs e)
        {
            if (EditTrackSht.Checked == true)
            {
                Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = true;
                TrkShtPane.Show();
            }
            else
            {
                Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = false;
                TrkShtPane.Hide();
            }
        }
        
        private void SOISelect_SelectionChanged(object sender, RibbonControlEventArgs e)
        {    
            if (RevSelect.Items.Count > 0)
            {
                string[] LTR = RevSelect.SelectedItem.Label.Split(' ');
                if (LTR[0] == "TBA")                                                                // Unprotect Doc or turn off rev tracking
                {
                    DialogResult result = MessageBox.Show("Save Changes to SOI?", "Save?",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                        SaveFile();

                    Globals.ThisDocument.TrackRevisions = false;
                }
                else
                {
                    Doc.UnProtect();
                }
            }
            else
            {
                Doc.UnProtect();
            }
            
            Word.Range Document = Globals.ThisDocument.Content;

            string[] LoadTitle = SOISelect.SelectedItem.Label.Split(' ');                       // Load Title of Selected SOI

            int i = 0;
            string Title = "";
            foreach (var text in LoadTitle)                                                     // Separate title from SOI #
            {
                if (i != 0)
                {
                    Title = Title + " " + text;
                }
                else
                    i++;
            }            

            RevSelect.Items.Clear();                                                                // Clear Revision LTR DropDown
            string[] memString = Globals.ThisDocument.RevSelect.Text.Split(',');                    // Prep to load Revision LTR DropDown             
            foreach(var item in memString)                                                      
            {
                if(item != "")
                {
                    string[] compare1 = item.Split(' ');
                    string[] compare2 = SOISelect.SelectedItem.Label.Split(' ');
                    int compareNums = compare1.Count();
                    if (compare1[compareNums - 1] == compare2[0])                                                 // If SOI # Matches, use Rev LTR in DropDown
                    {
                        RibbonDropDownItem ItemObj = Factory.CreateRibbonDropDownItem();
                        ItemObj.Label = compare1[compareNums - 2];
                        if(compareNums == 3)
                            ItemObj.Label = compare1[compareNums - 3] + " " + compare1[compareNums - 2];
                        RevSelect.Items.Add(ItemObj);                       
                    }                   
                }
            }

            string docName;
            if (RevSelect.Items.Count > 0)                                                          // If no revisions in dropdown, use blank SOI
            {
                RevSelect.SelectedItemIndex = RevSelect.Items.Count() - 1;                
                
                docName = Address + LoadTitle[0] + Title + @"\SOI-" + LoadTitle[0] + " " + "(" + RevSelect.SelectedItem.Label + ")" + Title + ".docx";
                
                Document.ImportFragment(docName);
                Globals.ThisDocument.HeaderTitle.Text = Globals.ThisDocument.Title.Text;
                Globals.ThisDocument.FooterLTR.Text = Globals.ThisDocument.LTR.Text;
                Globals.ThisDocument.FooterLTR2.Text = Globals.ThisDocument.LTR.Text;
                Globals.ThisDocument.FooterDateIssued.Text = Globals.ThisDocument.DateIssued.Text;
                Globals.ThisDocument.FooterDateIssued2.Text = Globals.ThisDocument.DateIssued.Text;
                Globals.ThisDocument.FooterDateRevised.Text = Globals.ThisDocument.DateRevised.Text;
                Globals.ThisDocument.FooterDateRevised2.Text = Globals.ThisDocument.DateRevised.Text;
                Globals.ThisDocument.FooterDateRevised.Text = Globals.ThisDocument.DateRevised.Text;
                Globals.ThisDocument.FooterSOINum.Text = LoadTitle[0];
                Globals.ThisDocument.FooterSOINum2.Text = LoadTitle[0];

                string[] LTR = RevSelect.SelectedItem.Label.Split(' ');

                if (LTR[0] == "TBA")                                                                // protect Doc or turn on rev tracking
                {
                    Globals.ThisDocument.TrackRevisions = true;
                    NewRev.Checked = true;
                }
                else
                {
                    Doc.Protect();
                    NewRev.Checked = false;
                }
            }
            else
            {
                docName = MasterAddress;                
                Document.ImportFragment(docName);

                string[] DocTitle = Title.Split(' ');
                Title = "";
                i = 0;
                foreach(var DocSplit in DocTitle)
                {
                    if(DocSplit != "")
                    {
                        if (i == 0)
                        {
                            Title = DocSplit;
                            i = 1;
                        }
                        else
                            Title = Title + " " + DocSplit;
                    }                    
                }

                Globals.ThisDocument.Title.Text = Title;
                Globals.ThisDocument.HeaderTitle.Text = Title;
                Globals.ThisDocument.FooterSOINum.Text = LoadTitle[0];
                Globals.ThisDocument.FooterSOINum2.Text = LoadTitle[0];
                Globals.ThisDocument.FooterDateIssued.Text = Globals.ThisDocument.DateIssued.Text;
                Globals.ThisDocument.FooterDateIssued2.Text = Globals.ThisDocument.DateIssued.Text;
                Globals.ThisDocument.FooterDateRevised.Text = Globals.ThisDocument.DateRevised.Text;
                Globals.ThisDocument.FooterDateRevised2.Text = Globals.ThisDocument.DateRevised.Text;                
                Globals.ThisDocument.FooterLTR.Text = Globals.ThisDocument.LTR.Text;
                Globals.ThisDocument.FooterLTR2.Text = Globals.ThisDocument.LTR.Text;

                NewRev.Checked = false;
                Doc.Protect();
            }               
        }

        private void NewSOI_Click(object sender, RibbonControlEventArgs e)
        {
            if (NewRev.Checked == false)
                Doc.UnProtect();
            else
                SaveFile();
                Globals.ThisDocument.TrackRevisions = false;
               
            Word.Range Document = Globals.ThisDocument.Content;            
            Document.ImportFragment(MasterAddress);
            RevSelect.Items.Clear();
            SOISelect.SelectedItemIndex = 0;            
            Globals.ThisDocument.HeaderTitle.Text = "";
            Globals.ThisDocument.Title.Text = "";
            Globals.ThisDocument.LTR.Text = "";
            Globals.ThisDocument.FooterLTR.Text = "";
            Globals.ThisDocument.FooterLTR2.Text = "";
            Globals.ThisDocument.Description.Text = "";
            Globals.ThisDocument.DateRevised.Text = "";
            Globals.ThisDocument.FooterDateRevised.Text = "";
            Globals.ThisDocument.FooterDateRevised2.Text = "";
            Globals.ThisDocument.DateIssued.Text = "";
            Globals.ThisDocument.FooterDateIssued.Text = "";
            Globals.ThisDocument.FooterDateIssued2.Text = "";
            Globals.ThisDocument.Auth.Text = "";
            Globals.ThisDocument.FooterSOINum.Text = "";
            Globals.ThisDocument.FooterSOINum2.Text = "";

            Doc.Protect();            
        }

        private void SaveFile()
        {
            Word.Range Document = Globals.ThisDocument.Content;
            string[] LTR = RevSelect.SelectedItem.Label.Split(' ');

            if (LTR[0] != "TBA")
                Doc.UnProtect();
            else
                Globals.ThisDocument.TrackRevisions = false;

            string docName = Address + Globals.ThisDocument.FooterSOINum.Text + " " + Globals.ThisDocument.Title.Text + @"\SOI-" + Globals.ThisDocument.FooterSOINum.Text + " " + "(" + Globals.ThisDocument.LTR.Text + ")" + " " + Globals.ThisDocument.Title.Text + ".docx";
            object fileName = docName;
            object missing = System.Reflection.Missing.Value;

            int max = Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Count;
            int i = 0;
            string labelCheck = Globals.ThisDocument.FooterSOINum.Text + " " + Globals.ThisDocument.Title.Text;
            bool newLabel = false;

            while (i < max)
            {
                if (labelCheck == Globals.Ribbons.StriteRevisionTab.SOISelect.Items[i].Label)
                {
                    newLabel = false;
                    i = max;
                }
                else
                {
                    newLabel = true;
                }
                i++;
            }

            if (newLabel == true)
            {
                DialogResult result = MessageBox.Show("Please make a new folder named: \n" + "SOI-" + Globals.ThisDocument.FooterSOINum.Text + " " + " " + Globals.ThisDocument.Title.Text + "\n \n at the following address: \n " + @"P:\ISO Documents\ISO MASTER SOI(AUTOMATED)" + "\n\n Or Cancel Submission", "Folder Not Found",
                MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (result == DialogResult.Retry)
                {
                    RibbonDropDownItem SOINewObj = Factory.CreateRibbonDropDownItem();
                    SOINewObj.Label = labelCheck;
                    SOISelect.Items.Add(SOINewObj);
                    SOISelect.SelectedItem = SOINewObj;
                }
                else
                {
                    MessageBox.Show("Submission Cancelled");
                    return;
                }                    
            }

            string[] compareString = Globals.ThisDocument.RevSelect.Text.Split(',');
            string storeString = Globals.ThisDocument.RevSelect.Text;
            bool dontStore = true;
            if (storeString == "")
            {
                dontStore = false;
            }
            else
            {
                foreach (var textString in compareString)
                {
                    if(textString != "")
                    {
                        if (textString == RevSelect.SelectedItem.Label + " " + Globals.ThisDocument.FooterSOINum.Text)
                            dontStore = true;
                        else
                            dontStore = false;
                    }                    
                }

            }

            if(dontStore == false)
            {
                storeString = storeString + "," + RevSelect.SelectedItem.Label + " " + Globals.ThisDocument.FooterSOINum.Text;
            }

            string[] SplitString = storeString.Split(',');
            string prevRev = "";
            string newStoreString = "";
            foreach (var rev in SplitString)
            {
                if(prevRev != "")
                {
                    string[] SplitPreRev = prevRev.Split(' ');
                    string[] SplitRev = rev.Split(' ');
                    if(SplitPreRev[1] != SplitRev[0])
                    {
                        newStoreString = newStoreString + "," + prevRev;
                    }
                }
                prevRev = rev;
            }
            newStoreString = newStoreString + "," + prevRev;
                        
            storeString = newStoreString;

            Globals.ThisDocument.RevSelect.Text = storeString;

            if (LTR[0] == "TBA")
                Globals.ThisDocument.TrackRevisions = true;
            else
                Doc.Protect();

            Globals.ThisDocument.SaveAs(ref fileName,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            if (LTR[0] != "TBA")
                Doc.UnProtect();
            else
                Globals.ThisDocument.TrackRevisions = false;

            Document.ImportFragment(MasterAddress);                                                       //Save SOI Select Data to MasterSOI Doc

            storeString = "";
            foreach (var item in SOISelect.Items)
            {
                storeString = storeString + "@" + item.Label;
            }

            Globals.ThisDocument.SOISelect.Text = storeString;
            Globals.ThisDocument.HeaderTitle.Text = "";
            Globals.ThisDocument.Title.Text = "";
            Globals.ThisDocument.LTR.Text = "";
            Globals.ThisDocument.FooterLTR.Text = "";
            Globals.ThisDocument.FooterLTR2.Text = "";
            Globals.ThisDocument.Description.Text = "";
            Globals.ThisDocument.DateRevised.Text = "";
            Globals.ThisDocument.FooterDateRevised.Text = "";
            Globals.ThisDocument.FooterDateRevised2.Text = "";
            Globals.ThisDocument.DateIssued.Text = "";
            Globals.ThisDocument.FooterDateIssued.Text = "";
            Globals.ThisDocument.FooterDateIssued2.Text = "";
            Globals.ThisDocument.Auth.Text = "";
            Globals.ThisDocument.FooterSOINum.Text = "";
            Globals.ThisDocument.FooterSOINum2.Text = "";

            Word.Selection wdAppSelection = Globals.ThisDocument.ActiveWindow.Selection;            
            object what = Word.WdGoToItem.wdGoToPage;
            object which = Word.WdGoToDirection.wdGoToAbsolute;
            object count = 2;                                                                                                           //change this number to specify the start of a different page
            Globals.ThisDocument.ActiveWindow.Selection.GoTo(ref what, ref which, ref count, ref missing);
            Object beginPageTwo = Globals.ThisDocument.ActiveWindow.Selection.Range.Start;                                              // This gets the start of the page specified by count object
            object end = Globals.ThisDocument.Content.End;
            Word.Range pg2 = Globals.ThisDocument.Range(ref beginPageTwo, ref end).FormattedText;                                                // modified this line per comments

            pg2.Text = "";

            fileName = MasterAddress;
            Globals.ThisDocument.SaveAs(ref fileName,
               ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing);

            Document = Globals.ThisDocument.Content;
            Document.ImportFragment(docName);

            string[] LoadTitle = SOISelect.SelectedItem.Label.Split(' ');                                                               // Load Title of Selected SOI

            Globals.ThisDocument.HeaderTitle.Text = Globals.ThisDocument.Title.Text;
            Globals.ThisDocument.FooterLTR.Text = Globals.ThisDocument.LTR.Text;
            Globals.ThisDocument.FooterLTR2.Text = Globals.ThisDocument.LTR.Text;
            Globals.ThisDocument.FooterDateIssued.Text = Globals.ThisDocument.DateIssued.Text;
            Globals.ThisDocument.FooterDateIssued2.Text = Globals.ThisDocument.DateIssued.Text;
            Globals.ThisDocument.FooterDateRevised.Text = Globals.ThisDocument.DateRevised.Text;
            Globals.ThisDocument.FooterDateRevised2.Text = Globals.ThisDocument.DateRevised.Text;
            Globals.ThisDocument.FooterDateRevised.Text = Globals.ThisDocument.DateRevised.Text;
            Globals.ThisDocument.FooterSOINum.Text = LoadTitle[0];
            Globals.ThisDocument.FooterSOINum2.Text = LoadTitle[0];                      

            if (LTR[0] == "TBA")
                Globals.ThisDocument.TrackRevisions = true;
            else
                Doc.Protect();
        }

        public void LoadFile()
        {
            int index = Globals.Ribbons.StriteRevisionTab.RevSelect.Items.Count - 1;
            int selectedItem = Globals.Ribbons.StriteRevisionTab.RevSelect.SelectedItemIndex;
            Word.Range Document = Globals.ThisDocument.Content;
            string docName;

            if (index > selectedItem)
            {
                Globals.Ribbons.StriteRevisionTab.RevSelect.SelectedItemIndex = index;                                                                           // Select last item in dropdown box
                string LTR = Globals.Ribbons.StriteRevisionTab.RevSelect.SelectedItem.Label;
                docName = Address + Globals.ThisDocument.FooterSOINum.Text + " " + Globals.ThisDocument.Title.Text + @"\SOI-" + Globals.ThisDocument.FooterSOINum.Text + " " + "(" + LTR + ")" + " " + Globals.ThisDocument.Title.Text + ".docx";
                Document.ImportFragment(docName);
                Globals.ThisDocument.FooterLTR.Text = Globals.ThisDocument.LTR.Text;
                Globals.ThisDocument.FooterLTR2.Text = Globals.ThisDocument.LTR.Text;
                Globals.ThisDocument.FooterDateRevised.Text = Globals.ThisDocument.DateRevised.Text;
                Globals.ThisDocument.FooterDateRevised2.Text = Globals.ThisDocument.DateRevised.Text;
            }
            
        }
        //////////////////////////////////////////////////////  OLD  //////////////////////////////////////////////////////////////
        //private void AcceptRev_Click(object sender, RibbonControlEventArgs e)
        //{
        //    if (NewRev.Checked == true)
        //    {
        //        object start = 0;
        //        object end = 0;

        //        Word.Range DocStart = Globals.ThisDocument.Range(ref start, ref end);               // Move Cursor to Start
        //        DocStart.Select();

        //        int RevNum = Globals.ThisDocument.Revisions.Count;                                  // Gets Number of Revisions
        //        int index = 0;
        //        int i = 0;

        //        string num = " ";
        //        string clauseNum = " ";

        //        List<string> checkNum = new List<string>();

        //        while (i < RevNum)
        //        {
        //            Globals.ThisDocument.ActiveWindow.Selection.NextRevision();
        //            Word.Range StartYRng = Globals.ThisDocument.ActiveWindow.Selection.Range;       // Gets Start Range for Revision Bar 

        //            Globals.ThisDocument.ActiveWindow.Selection.MoveDown();
        //            Word.Range EndYRng = Globals.ThisDocument.ActiveWindow.Selection.Range;         // Gets End Range for Revision Bar

        //            Word.Paragraphs para = Globals.ThisDocument.Paragraphs;

        //            int j = 1;

        //            while (j <= para.Count)
        //            {
        //                Word.Range parRange = Globals.ThisDocument.Paragraphs[j].Range;
        //                StartYRng.Select();
        //                bool InRange = Globals.ThisDocument.ActiveWindow.Selection.InRange(parRange);
        //                if (InRange == true)
        //                {
        //                    string parText = parRange.Text;

        //                    char num1 = parText[0];
        //                    char pt = parText[1];
        //                    char num2 = parText[2];

        //                    bool isNum1 = System.Char.IsDigit(num1);
        //                    bool isNum2 = System.Char.IsDigit(num2);

        //                    while (isNum1 == false || isNum2 == false)
        //                    {
        //                        j--;
        //                        parRange = Globals.ThisDocument.Paragraphs[j].Range;
        //                        parText = parRange.Text + "   ";
        //                        num1 = parText[1];
        //                        pt = parText[2];
        //                        num2 = parText[3];

        //                        isNum1 = System.Char.IsDigit(num1);
        //                        isNum2 = System.Char.IsDigit(num2);
        //                    }

        //                    num = $"{num1}" + $"{pt}" + $"{num2}";
        //                    checkNum.Add(num);

        //                    int listLength = checkNum.Count;

        //                    if (listLength > 1)
        //                    {
        //                        string prevNum = checkNum[index - 1];
        //                        if (num == prevNum)
        //                        { }
        //                        else
        //                        {
        //                            clauseNum = clauseNum + num + ", ";
        //                        }
        //                    }
        //                    else
        //                    {
        //                        clauseNum = clauseNum + num + ", ";
        //                    }
        //                    Globals.ThisDocument.ActiveWindow.Selection.NextRevision();
        //                    j = para.Count + 1;
        //                    index++;
        //                }
        //                else
        //                {
        //                    j++;
        //                }
        //            }

        //            float StartX = 40;
        //            float StartY = StartYRng.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
        //            float EndX = 40;
        //            float EndY = EndYRng.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];

        //            Globals.ThisDocument.TrackRevisions = false;
        //            Globals.ThisDocument.Shapes.AddLine(StartX, StartY, EndX, EndY);
        //            Globals.ThisDocument.TrackRevisions = true;

        //            i++;
        //        }

        //        Globals.ThisDocument.Description.Text = Globals.ThisDocument.Description.Text + "Revised Sections" + clauseNum;
        //        Globals.ThisDocument.AcceptAllRevisions();
        //    }
        //}
        //////////////////////////////////////////////////////  OLD  //////////////////////////////////////////////////////////////
    }
}
