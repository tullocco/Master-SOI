﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

#pragma warning disable 414
namespace Master_SOI {
    
    
    /// 
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class ThisDocument : Microsoft.Office.Tools.Word.DocumentBase {
        
        internal Microsoft.Office.Tools.ActionsPane ActionsPane;
        
        internal Microsoft.Office.Tools.Word.RichTextContentControl Title;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl LTR;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl Description;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl DateRevised;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl Auth;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl HeaderTitle;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterDateIssued2;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterSOINum2;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterLTR2;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterDateRevised2;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterDateIssued;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterSOINum;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterLTR;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl FooterDateRevised;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl SOISelect;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl RevSelect;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl IsProtec;
        
        internal Microsoft.Office.Tools.Word.PlainTextContentControl DateIssued;
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        internal Microsoft.Office.Interop.Word.Application ThisApplication;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public ThisDocument(global::Microsoft.Office.Tools.Word.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "ThisDocument", "ThisDocument") {
            Globals.Factory = factory;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            this.ThisApplication = this.GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
            Globals.ThisDocument = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            this.InternalStartup();
            this.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeDataBindings() {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeCachedData() {
            if ((this.DataHost == null)) {
                return;
            }
            if (this.DataHost.IsCacheInitialized) {
                this.DataHost.FillCachedData(this);
            }
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BindToData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StartCaching(string MemberName) {
            this.DataHost.StartCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StopCaching(string MemberName) {
            this.DataHost.StopCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool IsCached(string MemberName) {
            return this.DataHost.IsCached(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BeginInitialization() {
            this.BeginInit();
            this.ActionsPane.BeginInit();
            this.Title.BeginInit();
            this.LTR.BeginInit();
            this.Description.BeginInit();
            this.DateRevised.BeginInit();
            this.Auth.BeginInit();
            this.HeaderTitle.BeginInit();
            this.FooterDateIssued2.BeginInit();
            this.FooterSOINum2.BeginInit();
            this.FooterLTR2.BeginInit();
            this.FooterDateRevised2.BeginInit();
            this.FooterDateIssued.BeginInit();
            this.FooterSOINum.BeginInit();
            this.FooterLTR.BeginInit();
            this.FooterDateRevised.BeginInit();
            this.SOISelect.BeginInit();
            this.RevSelect.BeginInit();
            this.IsProtec.BeginInit();
            this.DateIssued.BeginInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void EndInitialization() {
            this.DateIssued.EndInit();
            this.IsProtec.EndInit();
            this.RevSelect.EndInit();
            this.SOISelect.EndInit();
            this.FooterDateRevised.EndInit();
            this.FooterLTR.EndInit();
            this.FooterSOINum.EndInit();
            this.FooterDateIssued.EndInit();
            this.FooterDateRevised2.EndInit();
            this.FooterLTR2.EndInit();
            this.FooterSOINum2.EndInit();
            this.FooterDateIssued2.EndInit();
            this.HeaderTitle.EndInit();
            this.Auth.EndInit();
            this.DateRevised.EndInit();
            this.Description.EndInit();
            this.LTR.EndInit();
            this.Title.EndInit();
            this.ActionsPane.EndInit();
            this.EndInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeControls() {
            this.ActionsPane = Globals.Factory.CreateActionsPane(null, null, "ActionsPane", "ActionsPane", this);
            this.Title = Globals.Factory.CreateRichTextContentControl(null, null, "3377258776", "Title", this);
            this.LTR = Globals.Factory.CreatePlainTextContentControl(null, null, "2427609775", "LTR", this);
            this.Description = Globals.Factory.CreatePlainTextContentControl(null, null, "876512493", "Description", this);
            this.DateRevised = Globals.Factory.CreatePlainTextContentControl(null, null, "3710579926", "DateRevised", this);
            this.Auth = Globals.Factory.CreatePlainTextContentControl(null, null, "416682546", "Auth", this);
            this.HeaderTitle = Globals.Factory.CreatePlainTextContentControl(null, null, "232438759", "HeaderTitle", this);
            this.FooterDateIssued2 = Globals.Factory.CreatePlainTextContentControl(null, null, "358326715", "FooterDateIssued2", this);
            this.FooterSOINum2 = Globals.Factory.CreatePlainTextContentControl(null, null, "3543545477", "FooterSOINum2", this);
            this.FooterLTR2 = Globals.Factory.CreatePlainTextContentControl(null, null, "3661177671", "FooterLTR2", this);
            this.FooterDateRevised2 = Globals.Factory.CreatePlainTextContentControl(null, null, "2803947009", "FooterDateRevised2", this);
            this.FooterDateIssued = Globals.Factory.CreatePlainTextContentControl(null, null, "3726926901", "FooterDateIssued", this);
            this.FooterSOINum = Globals.Factory.CreatePlainTextContentControl(null, null, "298730795", "FooterSOINum", this);
            this.FooterLTR = Globals.Factory.CreatePlainTextContentControl(null, null, "947280811", "FooterLTR", this);
            this.FooterDateRevised = Globals.Factory.CreatePlainTextContentControl(null, null, "1026370294", "FooterDateRevised", this);
            this.SOISelect = Globals.Factory.CreatePlainTextContentControl(null, null, "3898760545", "SOISelect", this);
            this.RevSelect = Globals.Factory.CreatePlainTextContentControl(null, null, "1481810242", "RevSelect", this);
            this.IsProtec = Globals.Factory.CreatePlainTextContentControl(null, null, "492529500", "IsProtec", this);
            this.DateIssued = Globals.Factory.CreatePlainTextContentControl(null, null, "3630654475", "DateIssued", this);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
            // 
            // ActionsPane
            // 
            this.ActionsPane.AutoSize = false;
            this.ActionsPane.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            // 
            // Title
            // 
            this.Title.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // LTR
            // 
            this.LTR.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // Description
            // 
            this.Description.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // DateRevised
            // 
            this.DateRevised.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // Auth
            // 
            this.Auth.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // HeaderTitle
            // 
            this.HeaderTitle.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterDateIssued2
            // 
            this.FooterDateIssued2.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterSOINum2
            // 
            this.FooterSOINum2.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterLTR2
            // 
            this.FooterLTR2.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterDateRevised2
            // 
            this.FooterDateRevised2.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterDateIssued
            // 
            this.FooterDateIssued.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterSOINum
            // 
            this.FooterSOINum.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterLTR
            // 
            this.FooterLTR.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // FooterDateRevised
            // 
            this.FooterDateRevised.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // SOISelect
            // 
            this.SOISelect.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // RevSelect
            // 
            this.RevSelect.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // IsProtec
            // 
            this.IsProtec.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // DateIssued
            // 
            this.DateIssued.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never;
            // 
            // ThisDocument
            // 
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void OnShutdown() {
            this.DateIssued.Dispose();
            this.IsProtec.Dispose();
            this.RevSelect.Dispose();
            this.SOISelect.Dispose();
            this.FooterDateRevised.Dispose();
            this.FooterLTR.Dispose();
            this.FooterSOINum.Dispose();
            this.FooterDateIssued.Dispose();
            this.FooterDateRevised2.Dispose();
            this.FooterLTR2.Dispose();
            this.FooterSOINum2.Dispose();
            this.FooterDateIssued2.Dispose();
            this.HeaderTitle.Dispose();
            this.Auth.Dispose();
            this.DateRevised.Dispose();
            this.Description.Dispose();
            this.LTR.Dispose();
            this.Title.Dispose();
            this.ActionsPane.Dispose();
            base.OnShutdown();
        }
    }
    
    /// 
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
    internal sealed partial class Globals {
        
        /// 
        private Globals() {
        }
        
        private static ThisDocument _ThisDocument;
        
        private static global::Microsoft.Office.Tools.Word.Factory _factory;
        
        private static ThisRibbonCollection _ThisRibbonCollection;
        
        internal static ThisDocument ThisDocument {
            get {
                return _ThisDocument;
            }
            set {
                if ((_ThisDocument == null)) {
                    _ThisDocument = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
        
        internal static global::Microsoft.Office.Tools.Word.Factory Factory {
            get {
                return _factory;
            }
            set {
                if ((_factory == null)) {
                    _factory = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
        
        internal static ThisRibbonCollection Ribbons {
            get {
                if ((_ThisRibbonCollection == null)) {
                    _ThisRibbonCollection = new ThisRibbonCollection(_factory.GetRibbonFactory());
                }
                return _ThisRibbonCollection;
            }
        }
    }
    
    /// 
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
    internal sealed partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonCollectionBase {
        
        /// 
        internal ThisRibbonCollection(global::Microsoft.Office.Tools.Ribbon.RibbonFactory factory) : 
                base(factory) {
        }
    }
}
