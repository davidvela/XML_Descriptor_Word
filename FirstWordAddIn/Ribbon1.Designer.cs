namespace FirstWordAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tabOthers = this.Factory.CreateRibbonTab();
            this.groupSaveAs = this.Factory.CreateRibbonGroup();
            this.buttonPDF = this.Factory.CreateRibbonButton();
            this.buttonXps = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.toggleButtonXMLDesc = this.Factory.CreateRibbonToggleButton();
            this.toggleButton_keepText = this.Factory.CreateRibbonToggleButton();
            this.buttonXMLDesc_All = this.Factory.CreateRibbonButton();
            this.button_InsertXMLD = this.Factory.CreateRibbonButton();
            this.button_cleanup_xmlD = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.tabOthers.SuspendLayout();
            this.groupSaveAs.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabOthers
            // 
            this.tabOthers.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabOthers.Groups.Add(this.groupSaveAs);
            this.tabOthers.Groups.Add(this.group1);
            this.tabOthers.Groups.Add(this.group2);
            this.tabOthers.Label = "Others Add_Ins";
            this.tabOthers.Name = "tabOthers";
            // 
            // groupSaveAs
            // 
            this.groupSaveAs.Items.Add(this.buttonPDF);
            this.groupSaveAs.Items.Add(this.buttonXps);
            this.groupSaveAs.Label = "Save As";
            this.groupSaveAs.Name = "groupSaveAs";
            // 
            // buttonPDF
            // 
            this.buttonPDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonPDF.Image = global::FirstWordAddIn.Properties.Resources.pdf_icon_48x48;
            this.buttonPDF.Label = "PDF";
            this.buttonPDF.Name = "buttonPDF";
            this.buttonPDF.ShowImage = true;
            this.buttonPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // buttonXps
            // 
            this.buttonXps.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonXps.Image = global::FirstWordAddIn.Properties.Resources.icon_xps;
            this.buttonXps.Label = "XPS";
            this.buttonXps.Name = "buttonXps";
            this.buttonXps.ShowImage = true;
            this.buttonXps.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonXps_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleButtonXMLDesc);
            this.group1.Items.Add(this.toggleButton_keepText);
            this.group1.Items.Add(this.buttonXMLDesc_All);
            this.group1.Items.Add(this.button_InsertXMLD);
            this.group1.Items.Add(this.button_cleanup_xmlD);
            this.group1.Label = "EasyQ";
            this.group1.Name = "group1";
            // 
            // toggleButtonXMLDesc
            // 
            this.toggleButtonXMLDesc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButtonXMLDesc.Image = global::FirstWordAddIn.Properties.Resources.xmlIcon;
            this.toggleButtonXMLDesc.Label = "AUTO XML Desription";
            this.toggleButtonXMLDesc.Name = "toggleButtonXMLDesc";
            this.toggleButtonXMLDesc.ShowImage = true;
            this.toggleButtonXMLDesc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonXMLDesc_Click);
            // 
            // toggleButton_keepText
            // 
            this.toggleButton_keepText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton_keepText.Image = global::FirstWordAddIn.Properties.Resources.keeptext;
            this.toggleButton_keepText.Label = "Keep Previous XML Description";
            this.toggleButton_keepText.Name = "toggleButton_keepText";
            this.toggleButton_keepText.ShowImage = true;
            this.toggleButton_keepText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton_keepText_Click);
            // 
            // buttonXMLDesc_All
            // 
            this.buttonXMLDesc_All.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonXMLDesc_All.Image = global::FirstWordAddIn.Properties.Resources.xml_128;
            this.buttonXMLDesc_All.Label = "Assign ALL XML Descriptions";
            this.buttonXMLDesc_All.Name = "buttonXMLDesc_All";
            this.buttonXMLDesc_All.ShowImage = true;
            this.buttonXMLDesc_All.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_XMLDesc_All);
            // 
            // button_InsertXMLD
            // 
            this.button_InsertXMLD.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_InsertXMLD.Image = global::FirstWordAddIn.Properties.Resources.download;
            this.button_InsertXMLD.Label = "Select to Insert XML Descriptions ";
            this.button_InsertXMLD.Name = "button_InsertXMLD";
            this.button_InsertXMLD.ShowImage = true;
            this.button_InsertXMLD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_InsertXMLD_Click);
            // 
            // button_cleanup_xmlD
            // 
            this.button_cleanup_xmlD.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_cleanup_xmlD.Image = global::FirstWordAddIn.Properties.Resources.images;
            this.button_cleanup_xmlD.Label = "Clean Up ALL XML Descriptions";
            this.button_cleanup_xmlD.Name = "button_cleanup_xmlD";
            this.button_cleanup_xmlD.ShowImage = true;
            this.button_cleanup_xmlD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_cleanup_xmlD_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.label1);
            this.group2.Label = "Information";
            this.group2.Name = "group2";
            // 
            // label1
            // 
            this.label1.Label = "Thanks for using this Add";
            this.label1.Name = "label1";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabOthers);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabOthers.ResumeLayout(false);
            this.tabOthers.PerformLayout();
            this.groupSaveAs.ResumeLayout(false);
            this.groupSaveAs.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabOthers;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSaveAs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonXMLDesc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonXps;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonXMLDesc_All;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_cleanup_xmlD;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_InsertXMLD;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton_keepText;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
