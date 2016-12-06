namespace ExcelAdjacencyListToDot
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
         this.group1 = this.Factory.CreateRibbonGroup();
         this.RunBtn = this.Factory.CreateRibbonButton();
         this.tab1.SuspendLayout();
         this.group1.SuspendLayout();
         // 
         // tab1
         // 
         this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
         this.tab1.Groups.Add(this.group1);
         this.tab1.Label = "Add Ins";
         this.tab1.Name = "tab1";
         // 
         // group1
         // 
         this.group1.Items.Add(this.RunBtn);
         this.group1.Label = "Dot";
         this.group1.Name = "group1";
         // 
         // RunBtn
         // 
         this.RunBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
         this.RunBtn.Image = global::ExcelAdjacencyListToDot.Properties.Resources.Icon;
         this.RunBtn.Label = "Process";
         this.RunBtn.Name = "RunBtn";
         this.RunBtn.ShowImage = true;
         this.RunBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RunBtn_Click);
         // 
         // Ribbon
         // 
         this.Name = "Ribbon";
         this.RibbonType = "Microsoft.Excel.Workbook";
         this.Tabs.Add(this.tab1);
         this.tab1.ResumeLayout(false);
         this.tab1.PerformLayout();
         this.group1.ResumeLayout(false);
         this.group1.PerformLayout();

      }

      #endregion

      internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
      internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
      internal Microsoft.Office.Tools.Ribbon.RibbonButton RunBtn;
   }

   partial class ThisRibbonCollection
   {
      internal Ribbon Ribbon1
      {
         get { return this.GetRibbon<Ribbon>(); }
      }
   }
}
