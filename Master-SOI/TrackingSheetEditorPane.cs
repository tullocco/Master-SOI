using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using System.Data;
using System.IO;

namespace Master_SOI
{
    public partial class TrackingSheetEditorPane : UserControl
    {
        public TrackingSheetEditorPane()
        {
            InitializeComponent();

            SubmitEdit.Click += SubmitEdit_Click;
            UpdateFields.Click += UpdateFields_Click;
            AddSign.Click += AddSign_Click;
        }

        private void AddSign_Click(object sender, EventArgs e)
        {
            object newSig = new object();
            newSig = NewSign.Text;
            if(NewSign.Text != "")
                checkedListBox1.Items.Add(newSig);
        }

        private void UpdateFields_Click(object sender, EventArgs e)
        {
            int i = 0;
            int j;
            int k;

            int max = checkedListBox1.Items.Count;

            while(i<max)
            {
                checkedListBox1.SetItemChecked(i, false);
                i++;
            }

            i = 0;
            while (i < max)
            {
                string compare = " " + checkedListBox1.Items[i].ToString();
                j = 0;

                while (j < checkedListBox1.Items.Count)
                {
                    string ApprovedBy = Globals.ThisDocument.Paragraphs[30 + j].Range.Text;

                    if (ApprovedBy == "Approved By:\r\a")
                    {
                        j = checkedListBox1.Items.Count;
                    }
                    else
                    {
                        char[] delimiterChars = { '_', ' ', '\r', '\a', '\t' };
                        string[] Signature = Globals.ThisDocument.Paragraphs[31+j].Range.Text.Split(delimiterChars);

                        string NewSignature = "";

                        k = 0;

                        foreach(var word in Signature)
                        {
                            if(word!= "")
                            {
                                NewSignature = NewSignature + " " + word;
                                k++;
                            }                            
                        }

                        if(compare == NewSignature)
                        {
                            checkedListBox1.SetItemChecked(i, true);
                            j = checkedListBox1.Items.Count;
                        }
                        else
                        {
                            j = j + 5;                           
                        }
                        
                    }
                }
                i++;
            }
           
            EditSOITitle.Text = Globals.ThisDocument.Title.Text;
            EditSOINum.Text = Globals.ThisDocument.FooterSOINum.Text;
            EditRevLTR.Text = Globals.ThisDocument.LTR.Text;
            EditDescription.Text = Globals.ThisDocument.Description.Text;
            EditAuth.Text = Globals.ThisDocument.Auth.Text;
            EditIssueDate.Text = Globals.ThisDocument.FooterDateIssued.Text;
        }

        private void SubmitEdit_Click(object sender, EventArgs e)
        {
            int i = 0;
            int max = checkedListBox1.CheckedItems.Count;
            bool unprotec;
            Word.Range Sign = Globals.ThisDocument.Paragraphs[27].Range;

            if (Globals.ThisDocument.TrackRevisions == true)
            {
                Globals.ThisDocument.TrackRevisions = false;
                unprotec = true;
            }                                               // Protection/TrackRevisions off
            else
            {
                Doc.UnProtect();
                unprotec = false;
            }
                
            while (i < checkedListBox1.Items.Count)
            {
                string ApprovedBy = Globals.ThisDocument.Paragraphs[30].Range.Text;

                if (ApprovedBy == "Approved By:\r\a")
                {
                    i = checkedListBox1.Items.Count;                    
                }
                else
                {
                    Globals.ThisDocument.Paragraphs[30].Range.Select();
                    Globals.ThisDocument.ActiveWindow.Selection.Rows.Delete();
                    i++;
                }
            }

            i = 0;

            if (max != 0)
            {
                while (i < max)
                {
                    Sign.Select();
                    Globals.ThisDocument.ActiveWindow.Selection.InsertRowsBelow();
                    string list = checkedListBox1.CheckedItems[i].ToString();
                    Globals.ThisDocument.Paragraphs[30].Range.Text = "Checked By: ";
                    Globals.ThisDocument.Paragraphs[31].Range.Text = "____________________  " + list;
                    Globals.ThisDocument.Paragraphs[32].Range.Text = "Date:";
                    Globals.ThisDocument.Paragraphs[33].Range.Text = " ______________";
                    i++;
                }
            }

            Globals.ThisDocument.Title.Text = EditSOITitle.Text;
            Globals.ThisDocument.HeaderTitle.Text = EditSOITitle.Text;
            Globals.ThisDocument.FooterSOINum.Text = EditSOINum.Text;
            Globals.ThisDocument.FooterSOINum2.Text = EditSOINum.Text;
            Globals.ThisDocument.LTR.Text = EditRevLTR.Text;
            Globals.ThisDocument.FooterLTR.Text = EditRevLTR.Text;
            Globals.ThisDocument.FooterLTR2.Text = EditRevLTR.Text;
            Globals.ThisDocument.Description.Text = EditDescription.Text;
            Globals.ThisDocument.Auth.Text = EditAuth.Text;
            Globals.ThisDocument.DateIssued.Text = EditIssueDate.Text;
            Globals.ThisDocument.FooterDateIssued.Text = EditIssueDate.Text;
            Globals.ThisDocument.FooterDateIssued2.Text = EditIssueDate.Text;

            if (unprotec == false)
            {
                Doc.Protect();
            }                                                                           // Protection/TrackRevisions on
            else
            {
                Globals.ThisDocument.TrackRevisions = true;
            }

            Globals.ThisDocument.Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneDocumentActions].Visible = false;
            Globals.Ribbons.StriteRevisionTab.EditTrackSht.Checked = false;
        }
        
    }


}
