using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace Master_SOI
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            //Doc.Protect();
            ActiveWindow.View.ReadingLayout = false;
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.FooterSOINum.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.FooterSOINum_Entering);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        private void FooterSOINum_Entering(object sender, ContentControlEnteringEventArgs e)
        {
        
        }
    }

    public class Doc
    {
        public static void Protect()
        {
            object noReset = false;
            object password = "";
            object useIRM = false;
            object enforceStyleLock = false;

            Globals.ThisDocument.Protect(Word.WdProtectionType.wdAllowOnlyReading,
                ref noReset, ref password, ref useIRM, ref enforceStyleLock);
        }

        public static void UnProtect()
        {
            object noReset = false;
            object password = "";
            object useIRM = false;
            object enforceStyleLock = false;

            Globals.ThisDocument.Unprotect(ref password);
        }

        public static void PopSOISelect()
        {
            string MasterAddress = @"P:\ISO Documents\ISO MASTER SOI(AUTOMATED)\Blank SOI-Master.docx";
            Word.Range Document = Globals.ThisDocument.Content;

            Document.ImportFragment(MasterAddress);
            Doc.Protect();

            if (Globals.ThisDocument.SOISelect.Text != "")
            {
                string[] SOISelect = Globals.ThisDocument.SOISelect.Text.Split('@');
                foreach(var Item in SOISelect)
                {
                    RibbonDropDownItem SOIAddObj = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                    SOIAddObj.Label = Item;
                    Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIAddObj);
                }
            }
            else
            {
                RibbonDropDownItem SOIObj0 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj1 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();  //Can't get array to work
                RibbonDropDownItem SOIObj2 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj3 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj4 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj5 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj6 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj9 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj10 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj11 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj13 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj15 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();               
                RibbonDropDownItem SOIObj19 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj20 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj22 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj23 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj25 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj26 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj27 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj28 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj29 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj30 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj31 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj32 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj33 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj34 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj35 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj36 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj37 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj38 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj39 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj42 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj43 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj45 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj48 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj49 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj50 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj51 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj52 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj53 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj54 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj55 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj59 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj60 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj63 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj64 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj70 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj71 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj72 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj78 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj79 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj83 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj87 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj88 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj93 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj95 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj96 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj98 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();                
                RibbonDropDownItem SOIObj100 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj102 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj104 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj105 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj106 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj107 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj108 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj109 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj110 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj111 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj114 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();
                RibbonDropDownItem SOIObj115 = Globals.Ribbons.StriteRevisionTab.Factory.CreateRibbonDropDownItem();

                SOIObj0.Label = "NEW SOI";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj0);

                SOIObj1.Label = "001 Customer Purchase Order Processing";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj1);

                SOIObj2.Label = "002 Control of Job Tickets";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj2);

                SOIObj3.Label = "003 Inspection & Test Reports";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj3);

                SOIObj4.Label = "004 Control of CNC Programs";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj4);

                SOIObj5.Label = "005 Control of Process Sheets";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj5);

                SOIObj6.Label = "006 Control of Standards and Specifications";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj6);

                SOIObj9.Label = "009 Control of Drawings";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj9);

                SOIObj10.Label = "010 Purchase Order Process & Quality Requirements";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj10);

                SOIObj11.Label = "011 Zeiss CMM Operation & Maintenance";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj11);

                SOIObj13.Label = "013 Internal Audit";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj13);

                SOIObj15.Label = "015 Equipment-Preventive Maintenance";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj15);

                SOIObj19.Label = "019 Helical Coil Installation";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj19);

                SOIObj20.Label = "020 Inspection and Test Plan (ITP)";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj20);

                SOIObj22.Label = "022 First Article Inspection";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj22);

                SOIObj23.Label = "023 Supplier Selection and Evaluation";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj23);

                SOIObj25.Label = "025 Final Product Release";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj25);

                SOIObj27.Label = "027 Shipping";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj27);

                SOIObj28.Label = "028 Context of the Organization";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj28);

                SOIObj29.Label = "029 Change Management";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj29);

                SOIObj31.Label = "031 Inventory";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj31);

                SOIObj33.Label = "033 Preservation of Product";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj33);

                SOIObj34.Label = "034 Raw Material";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj34);

                SOIObj36.Label = "036 Strite Nonconformance Reporting";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj36);

                SOIObj37.Label = "037 Corrective Action Report";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj37);

                SOIObj38.Label = "038 Customer Complaint + Return Authorization";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj38);

                SOIObj39.Label = "039 Advanced Quality Planning Level 1";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj39);

                SOIObj43.Label = "043 Quality Alert";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj43);

                SOIObj45.Label = "045 Supplier Auditing";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj45);

                SOIObj48.Label = "048 Calibration";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj48);

                SOIObj49.Label = "049 Control of Packing Slips";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj49);

                SOIObj51.Label = "051 Recieving Inspection";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj51);

                SOIObj52.Label = "052 Final Inspection";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj52);

                SOIObj53.Label = "053 Continual Improvement";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj53);

                SOIObj55.Label = "055 Quality Systems Awareness - Training Material";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj55);

                SOIObj59.Label = "059 Monitoring + Measuring and Process Performance Evaluation";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj59);

                SOIObj60.Label = "060 Gauge Control Log";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj60);

                SOIObj63.Label = "063 Basic Metrology";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj63);

                SOIObj64.Label = "064 Operator Requirements for Dimensional and Visual Inspection";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj64);

                SOIObj70.Label = "070 Designated Supplier Quality Assurance Representatives (DSQAR)";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj70);

                SOIObj71.Label = "071 Training";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj71);

                SOIObj72.Label = "072 Tool Room";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj72);

                SOIObj78.Label = "078 Controlled Goods Registration Program";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj78);

                SOIObj79.Label = "079 Validation of Equipment, Processes & Software";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj79);

                SOIObj83.Label = "083 Network Backups and Computer Lockouts";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj83);

                SOIObj87.Label = "087 First Off Inspection Process";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj87);

                SOIObj88.Label = "088 Control of Documents & Records";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj88);

                SOIObj93.Label = "093 Johnson Gauge";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj93);

                SOIObj95.Label = "095 Foreign Object Damage + Debris (FOD)";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj95);

                SOIObj96.Label = "096 Acceptance Authority Media Control";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj96);

                SOIObj98.Label = "098 Bushing + Sleeve + Bearing Installation";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj98);

                SOIObj100.Label = "100 Contingency Plan";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj100);

                SOIObj102.Label = "102 Risk Management Procedure";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj102);

                SOIObj104.Label = "104 Timely Reporting & Recall of Nonconforming Product";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj104);

                SOIObj105.Label = "105 Configuration Management";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj105);

                SOIObj106.Label = "106 Lee Plug Installation + Removal";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj106);

                SOIObj107.Label = "107 Management Review";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj107);

                SOIObj108.Label = "108 Infrastructure";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj108);

                SOIObj109.Label = "109 Inspection and Testing of Product";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj109);

                SOIObj110.Label = "110 Production Control";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj110);

                SOIObj111.Label = "111 Identification and Traceability";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj111);

                SOIObj114.Label = "114 Prevention and Detection of CFSI";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj114);

                SOIObj115.Label = "115 Security and Crisis Management";
                Globals.Ribbons.StriteRevisionTab.SOISelect.Items.Add(SOIObj115);
            }

        }
    }
}
