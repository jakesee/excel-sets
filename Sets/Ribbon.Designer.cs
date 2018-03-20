namespace Sets
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tabSets = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnNotIntersect = this.Factory.CreateRibbonButton();
            this.btnIntersection = this.Factory.CreateRibbonButton();
            this.btnNotSubset = this.Factory.CreateRibbonButton();
            this.btnUnion = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnFileMerge = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tabSets.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tabSets
            // 
            this.tabSets.Groups.Add(this.group1);
            this.tabSets.Groups.Add(this.group2);
            this.tabSets.Label = "Sets";
            this.tabSets.Name = "tabSets";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnNotIntersect);
            this.group1.Items.Add(this.btnIntersection);
            this.group1.Items.Add(this.btnNotSubset);
            this.group1.Items.Add(this.btnUnion);
            this.group1.Label = "Set Operations";
            this.group1.Name = "group1";
            // 
            // btnNotIntersect
            // 
            this.btnNotIntersect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNotIntersect.Image = global::Sets.Properties.Resources.venn_not_intersect;
            this.btnNotIntersect.Label = "Not Intersect";
            this.btnNotIntersect.Name = "btnNotIntersect";
            this.btnNotIntersect.ShowImage = true;
            this.btnNotIntersect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNotIntersect_Click);
            // 
            // btnIntersection
            // 
            this.btnIntersection.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnIntersection.Image = global::Sets.Properties.Resources.venn_intersect;
            this.btnIntersection.Label = "Intersect";
            this.btnIntersection.Name = "btnIntersection";
            this.btnIntersection.ShowImage = true;
            this.btnIntersection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnIntersection_Click);
            // 
            // btnNotSubset
            // 
            this.btnNotSubset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNotSubset.Image = global::Sets.Properties.Resources.venn_not_subset;
            this.btnNotSubset.Label = "Not Subset";
            this.btnNotSubset.Name = "btnNotSubset";
            this.btnNotSubset.ShowImage = true;
            this.btnNotSubset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNotSubset_Click);
            // 
            // btnUnion
            // 
            this.btnUnion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUnion.Image = global::Sets.Properties.Resources.venn_union;
            this.btnUnion.Label = "Union Distinct";
            this.btnUnion.Name = "btnUnion";
            this.btnUnion.ShowImage = true;
            this.btnUnion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnion_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnFileMerge);
            this.group2.Label = "File";
            this.group2.Name = "group2";
            // 
            // btnFileMerge
            // 
            this.btnFileMerge.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFileMerge.Label = "Merge";
            this.btnFileMerge.Name = "btnFileMerge";
            this.btnFileMerge.OfficeImageId = "MasterDocumentMergeSubdocuments";
            this.btnFileMerge.ShowImage = true;
            this.btnFileMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFileMerge_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabSets);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabSets.ResumeLayout(false);
            this.tabSets.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabSets;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNotIntersect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnIntersection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNotSubset;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnion;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFileMerge;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
