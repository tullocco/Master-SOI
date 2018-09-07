namespace Master_SOI
{
    partial class StriteRevisionTab : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public StriteRevisionTab()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StriteRevisionTab));
            this.StriteRevisions = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.SOISelect = this.Factory.CreateRibbonDropDown();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.RevSelect = this.Factory.CreateRibbonDropDown();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.NewSOI = this.Factory.CreateRibbonButton();
            this.EditTrackSht = this.Factory.CreateRibbonToggleButton();
            this.NewRev = this.Factory.CreateRibbonToggleButton();
            this.AcceptRev = this.Factory.CreateRibbonButton();
            this.RejectRev = this.Factory.CreateRibbonButton();
            this.Review = this.Factory.CreateRibbonButton();
            this.PrintRev = this.Factory.CreateRibbonButton();
            this.Submit = this.Factory.CreateRibbonButton();
            this.StriteRevisions.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // StriteRevisions
            // 
            this.StriteRevisions.Groups.Add(this.group1);
            this.StriteRevisions.Groups.Add(this.group2);
            this.StriteRevisions.Groups.Add(this.group3);
            this.StriteRevisions.Groups.Add(this.group4);
            this.StriteRevisions.Groups.Add(this.group5);
            this.StriteRevisions.Label = "STRITE REVISIONS";
            this.StriteRevisions.Name = "StriteRevisions";
            // 
            // group1
            // 
            this.group1.Items.Add(this.NewSOI);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.SOISelect);
            this.group1.Items.Add(this.EditTrackSht);
            this.group1.Label = "SELECT SOI";
            this.group1.Name = "group1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // SOISelect
            // 
            this.SOISelect.Label = "Select SOI";
            this.SOISelect.Name = "SOISelect";
            // 
            // group2
            // 
            this.group2.Items.Add(this.NewRev);
            this.group2.Items.Add(this.separator2);
            this.group2.Items.Add(this.RevSelect);
            this.group2.Label = "SELECT REVISION";
            this.group2.Name = "group2";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // RevSelect
            // 
            this.RevSelect.Label = "Select Revision";
            this.RevSelect.Name = "RevSelect";
            // 
            // group3
            // 
            this.group3.Items.Add(this.AcceptRev);
            this.group3.Items.Add(this.RejectRev);
            this.group3.Label = "CHANGES";
            this.group3.Name = "group3";
            // 
            // group4
            // 
            this.group4.Items.Add(this.Review);
            this.group4.Items.Add(this.PrintRev);
            this.group4.Label = "REVIEW CHANGES";
            this.group4.Name = "group4";
            // 
            // group5
            // 
            this.group5.Items.Add(this.Submit);
            this.group5.Label = "SUBMIT CHANGES";
            this.group5.Name = "group5";
            // 
            // NewSOI
            // 
            this.NewSOI.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.NewSOI.Image = ((System.Drawing.Image)(resources.GetObject("NewSOI.Image")));
            this.NewSOI.Label = "New SOI";
            this.NewSOI.Name = "NewSOI";
            this.NewSOI.ShowImage = true;
            // 
            // EditTrackSht
            // 
            this.EditTrackSht.Image = ((System.Drawing.Image)(resources.GetObject("EditTrackSht.Image")));
            this.EditTrackSht.Label = "Edit Tracking Sheet";
            this.EditTrackSht.Name = "EditTrackSht";
            this.EditTrackSht.ShowImage = true;
            // 
            // NewRev
            // 
            this.NewRev.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.NewRev.Image = ((System.Drawing.Image)(resources.GetObject("NewRev.Image")));
            this.NewRev.Label = "New Revision";
            this.NewRev.Name = "NewRev";
            this.NewRev.ShowImage = true;
            // 
            // AcceptRev
            // 
            this.AcceptRev.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AcceptRev.Image = ((System.Drawing.Image)(resources.GetObject("AcceptRev.Image")));
            this.AcceptRev.Label = "Accept Next Revision";
            this.AcceptRev.Name = "AcceptRev";
            this.AcceptRev.ShowImage = true;
            // 
            // RejectRev
            // 
            this.RejectRev.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RejectRev.Image = ((System.Drawing.Image)(resources.GetObject("RejectRev.Image")));
            this.RejectRev.Label = "Reject Next Revision";
            this.RejectRev.Name = "RejectRev";
            this.RejectRev.ShowImage = true;
            // 
            // Review
            // 
            this.Review.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Review.Image = ((System.Drawing.Image)(resources.GetObject("Review.Image")));
            this.Review.Label = "Review Submission";
            this.Review.Name = "Review";
            this.Review.ShowImage = true;
            // 
            // PrintRev
            // 
            this.PrintRev.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PrintRev.Image = ((System.Drawing.Image)(resources.GetObject("PrintRev.Image")));
            this.PrintRev.Label = "Print Review";
            this.PrintRev.Name = "PrintRev";
            this.PrintRev.ShowImage = true;
            // 
            // Submit
            // 
            this.Submit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Submit.Image = ((System.Drawing.Image)(resources.GetObject("Submit.Image")));
            this.Submit.Label = "Final Submission";
            this.Submit.Name = "Submit";
            this.Submit.ShowImage = true;
            // 
            // StriteRevisionTab
            // 
            this.Name = "StriteRevisionTab";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.StriteRevisions);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.StriteRevisionTab_Load);
            this.StriteRevisions.ResumeLayout(false);
            this.StriteRevisions.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab StriteRevisions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown SOISelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton NewRev;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown RevSelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AcceptRev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RejectRev;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Review;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Submit;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton EditTrackSht;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewSOI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PrintRev;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
    }

    partial class ThisRibbonCollection
    {
        internal StriteRevisionTab StriteRevisionTab
        {
            get { return this.GetRibbon<StriteRevisionTab>(); }
        }
    }
}
