namespace PowerPointAddIn1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.ChartQuickFormatColorsOnlyBtn = this.Factory.CreateRibbonButton();
            this.ChartQuickFormatFontsOnlyBtn = this.Factory.CreateRibbonButton();
            this.ChartQuickFormatNumberFOrmatsOnlyBtn = this.Factory.CreateRibbonButton();
            this.ChartQuickFormatPositionsOnlyBtn = this.Factory.CreateRibbonButton();
            this.ChartQuickAlignAnddSizeBtn = this.Factory.CreateRibbonButton();
            this.Tables = this.Factory.CreateRibbonGroup();
            this.TableFormatWithLayoutBtn = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.MakeSameSizeBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.Tables.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.Tables);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Fleur Addin";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.splitButton2);
            this.group1.Items.Add(this.ChartQuickAlignAnddSizeBtn);
            this.group1.Label = "Charts";
            this.group1.Name = "group1";
            // 
            // splitButton2
            // 
            this.splitButton2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton2.Image = ((System.Drawing.Image)(resources.GetObject("splitButton2.Image")));
            this.splitButton2.Items.Add(this.ChartQuickFormatColorsOnlyBtn);
            this.splitButton2.Items.Add(this.ChartQuickFormatFontsOnlyBtn);
            this.splitButton2.Items.Add(this.ChartQuickFormatNumberFOrmatsOnlyBtn);
            this.splitButton2.Items.Add(this.ChartQuickFormatPositionsOnlyBtn);
            this.splitButton2.Label = "Quick Format";
            this.splitButton2.Name = "splitButton2";
            // 
            // ChartQuickFormatColorsOnlyBtn
            // 
            this.ChartQuickFormatColorsOnlyBtn.Image = ((System.Drawing.Image)(resources.GetObject("ChartQuickFormatColorsOnlyBtn.Image")));
            this.ChartQuickFormatColorsOnlyBtn.Label = "Format Colors Only";
            this.ChartQuickFormatColorsOnlyBtn.Name = "ChartQuickFormatColorsOnlyBtn";
            this.ChartQuickFormatColorsOnlyBtn.ShowImage = true;
            this.ChartQuickFormatColorsOnlyBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChartQuickFormatColorsOnlyBtn_Click);
            // 
            // ChartQuickFormatFontsOnlyBtn
            // 
            this.ChartQuickFormatFontsOnlyBtn.Image = ((System.Drawing.Image)(resources.GetObject("ChartQuickFormatFontsOnlyBtn.Image")));
            this.ChartQuickFormatFontsOnlyBtn.Label = "Format Fonts Only";
            this.ChartQuickFormatFontsOnlyBtn.Name = "ChartQuickFormatFontsOnlyBtn";
            this.ChartQuickFormatFontsOnlyBtn.ShowImage = true;
            this.ChartQuickFormatFontsOnlyBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChartQuickFormatFontsOnlyBtn_Click);
            // 
            // ChartQuickFormatNumberFOrmatsOnlyBtn
            // 
            this.ChartQuickFormatNumberFOrmatsOnlyBtn.Image = ((System.Drawing.Image)(resources.GetObject("ChartQuickFormatNumberFOrmatsOnlyBtn.Image")));
            this.ChartQuickFormatNumberFOrmatsOnlyBtn.Label = "Format Numbers Formats Only";
            this.ChartQuickFormatNumberFOrmatsOnlyBtn.Name = "ChartQuickFormatNumberFOrmatsOnlyBtn";
            this.ChartQuickFormatNumberFOrmatsOnlyBtn.ShowImage = true;
            this.ChartQuickFormatNumberFOrmatsOnlyBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChartQuickFormatNumberFormatsOnlyBtn_Click);
            // 
            // ChartQuickFormatPositionsOnlyBtn
            // 
            this.ChartQuickFormatPositionsOnlyBtn.Image = ((System.Drawing.Image)(resources.GetObject("ChartQuickFormatPositionsOnlyBtn.Image")));
            this.ChartQuickFormatPositionsOnlyBtn.Label = "Format Positions Only";
            this.ChartQuickFormatPositionsOnlyBtn.Name = "ChartQuickFormatPositionsOnlyBtn";
            this.ChartQuickFormatPositionsOnlyBtn.ShowImage = true;
            this.ChartQuickFormatPositionsOnlyBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChartQuickFormatPositionsOnlyBtn_Click);
            // 
            // ChartQuickAlignAnddSizeBtn
            // 
            this.ChartQuickAlignAnddSizeBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ChartQuickAlignAnddSizeBtn.Image = ((System.Drawing.Image)(resources.GetObject("ChartQuickAlignAnddSizeBtn.Image")));
            this.ChartQuickAlignAnddSizeBtn.Label = "Quick Align and Size";
            this.ChartQuickAlignAnddSizeBtn.Name = "ChartQuickAlignAnddSizeBtn";
            this.ChartQuickAlignAnddSizeBtn.ShowImage = true;
            this.ChartQuickAlignAnddSizeBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChartQuickAlignAndSizeBtn_Click);
            // 
            // Tables
            // 
            this.Tables.Items.Add(this.TableFormatWithLayoutBtn);
            this.Tables.Label = "Tables";
            this.Tables.Name = "Tables";
            // 
            // TableFormatWithLayoutBtn
            // 
            this.TableFormatWithLayoutBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.TableFormatWithLayoutBtn.Image = ((System.Drawing.Image)(resources.GetObject("TableFormatWithLayoutBtn.Image")));
            this.TableFormatWithLayoutBtn.Label = "Format table in template style";
            this.TableFormatWithLayoutBtn.Name = "TableFormatWithLayoutBtn";
            this.TableFormatWithLayoutBtn.ShowImage = true;
            this.TableFormatWithLayoutBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableFormatWithLayoutBtn_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.MakeSameSizeBtn);
            this.group2.Label = "Position and Size";
            this.group2.Name = "group2";
            // 
            // MakeSameSizeBtn
            // 
            this.MakeSameSizeBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MakeSameSizeBtn.Image = ((System.Drawing.Image)(resources.GetObject("MakeSameSizeBtn.Image")));
            this.MakeSameSizeBtn.Label = "Make Same Size";
            this.MakeSameSizeBtn.Name = "MakeSameSizeBtn";
            this.MakeSameSizeBtn.ShowImage = true;
            this.MakeSameSizeBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MakeSameSizeBtn_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Tables.ResumeLayout(false);
            this.Tables.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChartQuickFormatColorsOnlyBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChartQuickFormatFontsOnlyBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChartQuickFormatNumberFOrmatsOnlyBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChartQuickFormatPositionsOnlyBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChartQuickAlignAnddSizeBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Tables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TableFormatWithLayoutBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MakeSameSizeBtn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
