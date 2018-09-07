namespace Master_SOI
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class TrackingSheetEditorPane
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
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
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.EditSOITitle = new System.Windows.Forms.TextBox();
            this.EditSOITitleLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.EditSOINum = new System.Windows.Forms.TextBox();
            this.EditRevLTR = new System.Windows.Forms.TextBox();
            this.EditDescription = new System.Windows.Forms.TextBox();
            this.EditAuth = new System.Windows.Forms.TextBox();
            this.EditIssueDate = new System.Windows.Forms.TextBox();
            this.SubmitEdit = new System.Windows.Forms.Button();
            this.UpdateFields = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.NewSign = new System.Windows.Forms.TextBox();
            this.AddSign = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Items.AddRange(new object[] {
            "Wire EDM Supv.",
            "Training Coord.",
            "Tool and Cutter",
            "Surface Grinding Supv.",
            "Source Insp.",
            "Receiving Insp.",
            "Quality Tech.",
            "QMR",
            "QC Manager",
            "QA Manager",
            "Purchasing",
            "Production Manager",
            "Process Eng.",
            "Plating Supv.",
            "Plant Superintendent",
            "Plant Manager",
            "NC Lathes Supv.",
            "Milling Supv.",
            "Marketing Manager",
            "Maintenance",
            "Long Manf. Supv.",
            "Lapping Supv.",
            "JHSC",
            "Inspection Supv.",
            "Human Resources",
            "Hor. Mills Supv.",
            "Front Office",
            "Eng. Supv.",
            "Eng. Rep.",
            "Eng. Manager",
            "Eng. Contract Admin.",
            "Eng. Assistant",
            "Die Sink EDM",
            "Development Supv.",
            "Deburring",
            "Cylindrical Grinding Supv.",
            "Cust. Doc. Prep.",
            "Controller",
            "CHNC Supv.",
            "Calibration Tech.",
            "Assembly Supv.",
            "Admin. Office"});
            this.checkedListBox1.Location = new System.Drawing.Point(309, 34);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(154, 499);
            this.checkedListBox1.TabIndex = 0;
            // 
            // EditSOITitle
            // 
            this.EditSOITitle.Location = new System.Drawing.Point(98, 34);
            this.EditSOITitle.Name = "EditSOITitle";
            this.EditSOITitle.Size = new System.Drawing.Size(191, 20);
            this.EditSOITitle.TabIndex = 1;
            // 
            // EditSOITitleLabel
            // 
            this.EditSOITitleLabel.AutoSize = true;
            this.EditSOITitleLabel.Location = new System.Drawing.Point(16, 37);
            this.EditSOITitleLabel.Name = "EditSOITitleLabel";
            this.EditSOITitleLabel.Size = new System.Drawing.Size(48, 13);
            this.EditSOITitleLabel.TabIndex = 2;
            this.EditSOITitleLabel.Text = "SOI Title";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 113);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Revision LTR";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(306, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Required Signatures";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 156);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Description";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 77);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "SOI Number";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 264);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 13);
            this.label5.TabIndex = 7;
            this.label5.Text = "Issue Date";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(16, 228);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 13);
            this.label6.TabIndex = 8;
            this.label6.Text = "Authorization";
            // 
            // EditSOINum
            // 
            this.EditSOINum.Location = new System.Drawing.Point(98, 74);
            this.EditSOINum.Name = "EditSOINum";
            this.EditSOINum.Size = new System.Drawing.Size(45, 20);
            this.EditSOINum.TabIndex = 9;
            // 
            // EditRevLTR
            // 
            this.EditRevLTR.Location = new System.Drawing.Point(98, 110);
            this.EditRevLTR.Name = "EditRevLTR";
            this.EditRevLTR.Size = new System.Drawing.Size(45, 20);
            this.EditRevLTR.TabIndex = 10;
            // 
            // EditDescription
            // 
            this.EditDescription.Location = new System.Drawing.Point(98, 153);
            this.EditDescription.Multiline = true;
            this.EditDescription.Name = "EditDescription";
            this.EditDescription.Size = new System.Drawing.Size(191, 60);
            this.EditDescription.TabIndex = 11;
            // 
            // EditAuth
            // 
            this.EditAuth.Location = new System.Drawing.Point(98, 225);
            this.EditAuth.Name = "EditAuth";
            this.EditAuth.Size = new System.Drawing.Size(45, 20);
            this.EditAuth.TabIndex = 12;
            // 
            // EditIssueDate
            // 
            this.EditIssueDate.Location = new System.Drawing.Point(98, 261);
            this.EditIssueDate.Name = "EditIssueDate";
            this.EditIssueDate.Size = new System.Drawing.Size(81, 20);
            this.EditIssueDate.TabIndex = 13;
            // 
            // SubmitEdit
            // 
            this.SubmitEdit.Location = new System.Drawing.Point(309, 581);
            this.SubmitEdit.Name = "SubmitEdit";
            this.SubmitEdit.Size = new System.Drawing.Size(154, 63);
            this.SubmitEdit.TabIndex = 14;
            this.SubmitEdit.Text = "Submit Changes";
            this.SubmitEdit.UseVisualStyleBackColor = true;
            // 
            // UpdateFields
            // 
            this.UpdateFields.Location = new System.Drawing.Point(204, 252);
            this.UpdateFields.Name = "UpdateFields";
            this.UpdateFields.Size = new System.Drawing.Size(85, 29);
            this.UpdateFields.TabIndex = 15;
            this.UpdateFields.Text = "Update Fields";
            this.UpdateFields.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 519);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(77, 13);
            this.label7.TabIndex = 16;
            this.label7.Text = "New Signature";
            // 
            // NewSign
            // 
            this.NewSign.Location = new System.Drawing.Point(98, 512);
            this.NewSign.Name = "NewSign";
            this.NewSign.Size = new System.Drawing.Size(132, 20);
            this.NewSign.TabIndex = 17;
            // 
            // AddSign
            // 
            this.AddSign.Location = new System.Drawing.Point(250, 509);
            this.AddSign.Name = "AddSign";
            this.AddSign.Size = new System.Drawing.Size(39, 23);
            this.AddSign.TabIndex = 18;
            this.AddSign.Text = "Add";
            this.AddSign.UseVisualStyleBackColor = true;
            // 
            // TrackingSheetEditorPane
            // 
            this.AutoSize = true;
            this.Controls.Add(this.AddSign);
            this.Controls.Add(this.NewSign);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.UpdateFields);
            this.Controls.Add(this.SubmitEdit);
            this.Controls.Add(this.EditIssueDate);
            this.Controls.Add(this.EditAuth);
            this.Controls.Add(this.EditDescription);
            this.Controls.Add(this.EditRevLTR);
            this.Controls.Add(this.EditSOINum);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.EditSOITitleLabel);
            this.Controls.Add(this.EditSOITitle);
            this.Controls.Add(this.checkedListBox1);
            this.Name = "TrackingSheetEditorPane";
            this.Size = new System.Drawing.Size(678, 678);
            this.Load += new System.EventHandler(this.UpdateFields_Click);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.TextBox EditSOITitle;
        private System.Windows.Forms.Label EditSOITitleLabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox EditSOINum;
        private System.Windows.Forms.TextBox EditRevLTR;
        private System.Windows.Forms.TextBox EditDescription;
        private System.Windows.Forms.TextBox EditAuth;
        private System.Windows.Forms.TextBox EditIssueDate;
        private System.Windows.Forms.Button SubmitEdit;
        private System.Windows.Forms.Button UpdateFields;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox NewSign;
        private System.Windows.Forms.Button AddSign;
    }
}
