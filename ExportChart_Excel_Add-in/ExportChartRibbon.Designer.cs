namespace ExportChart_Excel_Add_in
{
    partial class ExportChartRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExportChartRibbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.exportGroup = this.Factory.CreateRibbonGroup();
            this.helpGroup = this.Factory.CreateRibbonGroup();
            this.pdfBtn = this.Factory.CreateRibbonButton();
            this.imgBtn = this.Factory.CreateRibbonButton();
            this.aboutBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.exportGroup.SuspendLayout();
            this.helpGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.exportGroup);
            this.tab1.Groups.Add(this.helpGroup);
            this.tab1.Label = "Export Chart";
            this.tab1.Name = "tab1";
            // 
            // exportGroup
            // 
            this.exportGroup.Items.Add(this.pdfBtn);
            this.exportGroup.Items.Add(this.imgBtn);
            this.exportGroup.Label = "Export";
            this.exportGroup.Name = "exportGroup";
            // 
            // helpGroup
            // 
            this.helpGroup.Items.Add(this.aboutBtn);
            this.helpGroup.Label = "Help";
            this.helpGroup.Name = "helpGroup";
            // 
            // pdfBtn
            // 
            this.pdfBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.pdfBtn.Description = "Export selected chart to PDF";
            this.pdfBtn.Image = global::ExportChart_Excel_Add_in.Properties.Resources.PDF;
            this.pdfBtn.Label = "Export to PDF";
            this.pdfBtn.Name = "pdfBtn";
            this.pdfBtn.ShowImage = true;
            this.pdfBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pdfBtn_Click);
            // 
            // imgBtn
            // 
            this.imgBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.imgBtn.Description = "Export selected chart to image";
            this.imgBtn.Image = global::ExportChart_Excel_Add_in.Properties.Resources.IMG;
            this.imgBtn.Label = "Export to Image";
            this.imgBtn.Name = "imgBtn";
            this.imgBtn.ShowImage = true;
            this.imgBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.imgBtn_Click);
            // 
            // aboutBtn
            // 
            this.aboutBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.aboutBtn.Description = "About this add-in";
            this.aboutBtn.Image = global::ExportChart_Excel_Add_in.Properties.Resources.About;
            this.aboutBtn.Label = "About";
            this.aboutBtn.Name = "aboutBtn";
            this.aboutBtn.ShowImage = true;
            this.aboutBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.aboutBtn_Click);
            // 
            // ExportChartRibbon
            // 
            this.Name = "ExportChartRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExportChartRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.exportGroup.ResumeLayout(false);
            this.exportGroup.PerformLayout();
            this.helpGroup.ResumeLayout(false);
            this.helpGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup exportGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pdfBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton imgBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup helpGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton aboutBtn;
    }

    partial class ThisRibbonCollection
    {
        internal ExportChartRibbon Ribbon1
        {
            get { return this.GetRibbon<ExportChartRibbon>(); }
        }
    }
}
