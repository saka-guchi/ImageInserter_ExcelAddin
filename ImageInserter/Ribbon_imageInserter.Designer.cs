
namespace ImageInserter
{
    partial class Ribbon_imageInserter : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_imageInserter()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon_imageInserter));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            this.tab_imageInserter = this.Factory.CreateRibbonTab();
            this.group_insert = this.Factory.CreateRibbonGroup();
            this.splitButton_insert = this.Factory.CreateRibbonSplitButton();
            this.button_insertFile = this.Factory.CreateRibbonButton();
            this.button_insertLink = this.Factory.CreateRibbonButton();
            this.button_insertFolder = this.Factory.CreateRibbonButton();
            this.group_setting = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.checkBox_cell = this.Factory.CreateRibbonCheckBox();
            this.checkBox_memo = this.Factory.CreateRibbonCheckBox();
            this.group_cell = this.Factory.CreateRibbonGroup();
            this.checkBox_setSize = this.Factory.CreateRibbonCheckBox();
            this.editBox_setW = this.Factory.CreateRibbonEditBox();
            this.editBox_setH = this.Factory.CreateRibbonEditBox();
            this.dropDown_shrink = this.Factory.CreateRibbonDropDown();
            this.dropDown_writeCell = this.Factory.CreateRibbonDropDown();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.dropDown_direction = this.Factory.CreateRibbonDropDown();
            this.group_memo = this.Factory.CreateRibbonGroup();
            this.checkBox_maxSize = this.Factory.CreateRibbonCheckBox();
            this.editBox_maxW = this.Factory.CreateRibbonEditBox();
            this.editBox_maxH = this.Factory.CreateRibbonEditBox();
            this.dropDown_writeMemo = this.Factory.CreateRibbonDropDown();
            this.tab_imageInserter.SuspendLayout();
            this.group_insert.SuspendLayout();
            this.group_setting.SuspendLayout();
            this.group_cell.SuspendLayout();
            this.group_memo.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_imageInserter
            // 
            this.tab_imageInserter.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_imageInserter.Groups.Add(this.group_insert);
            this.tab_imageInserter.Groups.Add(this.group_setting);
            this.tab_imageInserter.Groups.Add(this.group_cell);
            this.tab_imageInserter.Groups.Add(this.group_memo);
            resources.ApplyResources(this.tab_imageInserter, "tab_imageInserter");
            this.tab_imageInserter.Name = "tab_imageInserter";
            // 
            // group_insert
            // 
            this.group_insert.Items.Add(this.splitButton_insert);
            resources.ApplyResources(this.group_insert, "group_insert");
            this.group_insert.Name = "group_insert";
            // 
            // splitButton_insert
            // 
            this.splitButton_insert.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.splitButton_insert, "splitButton_insert");
            this.splitButton_insert.Items.Add(this.button_insertFile);
            this.splitButton_insert.Items.Add(this.button_insertLink);
            this.splitButton_insert.Items.Add(this.button_insertFolder);
            this.splitButton_insert.Name = "splitButton_insert";
            this.splitButton_insert.OfficeImageId = "RestoreImageSize";
            this.splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertFile_Click);
            // 
            // button_insertFile
            // 
            resources.ApplyResources(this.button_insertFile, "button_insertFile");
            this.button_insertFile.Name = "button_insertFile";
            this.button_insertFile.OfficeImageId = "RestoreImageSize";
            this.button_insertFile.ShowImage = true;
            this.button_insertFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertFile_Click);
            // 
            // button_insertLink
            // 
            resources.ApplyResources(this.button_insertLink, "button_insertLink");
            this.button_insertLink.Name = "button_insertLink";
            this.button_insertLink.OfficeImageId = "OmsImageFromClip";
            this.button_insertLink.ShowImage = true;
            this.button_insertLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertLink_Click);
            // 
            // button_insertFolder
            // 
            resources.ApplyResources(this.button_insertFolder, "button_insertFolder");
            this.button_insertFolder.Name = "button_insertFolder";
            this.button_insertFolder.OfficeImageId = "ApplyImageBackgroundTile";
            this.button_insertFolder.ShowImage = true;
            this.button_insertFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertFolder_Click);
            // 
            // group_setting
            // 
            this.group_setting.Items.Add(this.label1);
            this.group_setting.Items.Add(this.checkBox_cell);
            this.group_setting.Items.Add(this.checkBox_memo);
            resources.ApplyResources(this.group_setting, "group_setting");
            this.group_setting.Name = "group_setting";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // checkBox_cell
            // 
            this.checkBox_cell.Checked = true;
            resources.ApplyResources(this.checkBox_cell, "checkBox_cell");
            this.checkBox_cell.Name = "checkBox_cell";
            // 
            // checkBox_memo
            // 
            this.checkBox_memo.Checked = true;
            resources.ApplyResources(this.checkBox_memo, "checkBox_memo");
            this.checkBox_memo.Name = "checkBox_memo";
            // 
            // group_cell
            // 
            this.group_cell.Items.Add(this.checkBox_setSize);
            this.group_cell.Items.Add(this.editBox_setW);
            this.group_cell.Items.Add(this.editBox_setH);
            this.group_cell.Items.Add(this.dropDown_shrink);
            this.group_cell.Items.Add(this.dropDown_writeCell);
            this.group_cell.Items.Add(this.separator2);
            this.group_cell.Items.Add(this.label2);
            this.group_cell.Items.Add(this.dropDown_direction);
            resources.ApplyResources(this.group_cell, "group_cell");
            this.group_cell.Name = "group_cell";
            // 
            // checkBox_setSize
            // 
            resources.ApplyResources(this.checkBox_setSize, "checkBox_setSize");
            this.checkBox_setSize.Name = "checkBox_setSize";
            this.checkBox_setSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_setSize_Click);
            // 
            // editBox_setW
            // 
            resources.ApplyResources(this.editBox_setW, "editBox_setW");
            this.editBox_setW.MaxLength = 4;
            this.editBox_setW.Name = "editBox_setW";
            // 
            // editBox_setH
            // 
            resources.ApplyResources(this.editBox_setH, "editBox_setH");
            this.editBox_setH.MaxLength = 4;
            this.editBox_setH.Name = "editBox_setH";
            // 
            // dropDown_shrink
            // 
            resources.ApplyResources(this.dropDown_shrink, "dropDown_shrink");
            resources.ApplyResources(ribbonDropDownItemImpl1, "ribbonDropDownItemImpl1");
            ribbonDropDownItemImpl1.OfficeImageId = "BackgroundImageGallery";
            ribbonDropDownItemImpl1.Tag = "fit";
            resources.ApplyResources(ribbonDropDownItemImpl2, "ribbonDropDownItemImpl2");
            ribbonDropDownItemImpl2.OfficeImageId = "CellHeight";
            ribbonDropDownItemImpl2.Tag = "fitW";
            resources.ApplyResources(ribbonDropDownItemImpl3, "ribbonDropDownItemImpl3");
            ribbonDropDownItemImpl3.OfficeImageId = "GroupTableCellFormat";
            ribbonDropDownItemImpl3.Tag = "fitH";
            this.dropDown_shrink.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown_shrink.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown_shrink.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown_shrink.Name = "dropDown_shrink";
            this.dropDown_shrink.OfficeImageId = "DiagramResizeClassic";
            this.dropDown_shrink.ShowImage = true;
            this.dropDown_shrink.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_shrink_SelectionChanged);
            // 
            // dropDown_writeCell
            // 
            resources.ApplyResources(this.dropDown_writeCell, "dropDown_writeCell");
            resources.ApplyResources(ribbonDropDownItemImpl4, "ribbonDropDownItemImpl4");
            ribbonDropDownItemImpl4.OfficeImageId = "CancelRequest";
            ribbonDropDownItemImpl4.Tag = "none";
            resources.ApplyResources(ribbonDropDownItemImpl5, "ribbonDropDownItemImpl5");
            ribbonDropDownItemImpl5.OfficeImageId = "FileNew";
            ribbonDropDownItemImpl5.Tag = "name";
            resources.ApplyResources(ribbonDropDownItemImpl6, "ribbonDropDownItemImpl6");
            ribbonDropDownItemImpl6.OfficeImageId = "GroupImapFolderOptions";
            ribbonDropDownItemImpl6.Tag = "path";
            this.dropDown_writeCell.Items.Add(ribbonDropDownItemImpl4);
            this.dropDown_writeCell.Items.Add(ribbonDropDownItemImpl5);
            this.dropDown_writeCell.Items.Add(ribbonDropDownItemImpl6);
            this.dropDown_writeCell.Name = "dropDown_writeCell";
            this.dropDown_writeCell.OfficeImageId = "IconPencilTool";
            this.dropDown_writeCell.ShowImage = true;
            this.dropDown_writeCell.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_shrink_SelectionChanged);
            // 
            // separator2
            // 
            resources.ApplyResources(this.separator2, "separator2");
            this.separator2.Name = "separator2";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // dropDown_direction
            // 
            resources.ApplyResources(this.dropDown_direction, "dropDown_direction");
            resources.ApplyResources(ribbonDropDownItemImpl7, "ribbonDropDownItemImpl7");
            ribbonDropDownItemImpl7.OfficeImageId = "ChartNavDrillDown";
            ribbonDropDownItemImpl7.Tag = "under";
            resources.ApplyResources(ribbonDropDownItemImpl8, "ribbonDropDownItemImpl8");
            ribbonDropDownItemImpl8.OfficeImageId = "OrgChartReportMoveRight";
            ribbonDropDownItemImpl8.Tag = "right";
            this.dropDown_direction.Items.Add(ribbonDropDownItemImpl7);
            this.dropDown_direction.Items.Add(ribbonDropDownItemImpl8);
            this.dropDown_direction.Name = "dropDown_direction";
            // 
            // group_memo
            // 
            this.group_memo.Items.Add(this.checkBox_maxSize);
            this.group_memo.Items.Add(this.editBox_maxW);
            this.group_memo.Items.Add(this.editBox_maxH);
            this.group_memo.Items.Add(this.dropDown_writeMemo);
            resources.ApplyResources(this.group_memo, "group_memo");
            this.group_memo.Name = "group_memo";
            // 
            // checkBox_maxSize
            // 
            this.checkBox_maxSize.Checked = true;
            resources.ApplyResources(this.checkBox_maxSize, "checkBox_maxSize");
            this.checkBox_maxSize.Name = "checkBox_maxSize";
            this.checkBox_maxSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_maxSize_Click);
            // 
            // editBox_maxW
            // 
            resources.ApplyResources(this.editBox_maxW, "editBox_maxW");
            this.editBox_maxW.MaxLength = 4;
            this.editBox_maxW.Name = "editBox_maxW";
            // 
            // editBox_maxH
            // 
            resources.ApplyResources(this.editBox_maxH, "editBox_maxH");
            this.editBox_maxH.MaxLength = 4;
            this.editBox_maxH.Name = "editBox_maxH";
            // 
            // dropDown_writeMemo
            // 
            resources.ApplyResources(this.dropDown_writeMemo, "dropDown_writeMemo");
            resources.ApplyResources(ribbonDropDownItemImpl9, "ribbonDropDownItemImpl9");
            ribbonDropDownItemImpl9.OfficeImageId = "CancelRequest";
            ribbonDropDownItemImpl9.Tag = "none";
            resources.ApplyResources(ribbonDropDownItemImpl10, "ribbonDropDownItemImpl10");
            ribbonDropDownItemImpl10.OfficeImageId = "FileNew";
            ribbonDropDownItemImpl10.Tag = "name";
            resources.ApplyResources(ribbonDropDownItemImpl11, "ribbonDropDownItemImpl11");
            ribbonDropDownItemImpl11.OfficeImageId = "GroupImapFolderOptions";
            ribbonDropDownItemImpl11.Tag = "path";
            this.dropDown_writeMemo.Items.Add(ribbonDropDownItemImpl9);
            this.dropDown_writeMemo.Items.Add(ribbonDropDownItemImpl10);
            this.dropDown_writeMemo.Items.Add(ribbonDropDownItemImpl11);
            this.dropDown_writeMemo.Name = "dropDown_writeMemo";
            this.dropDown_writeMemo.OfficeImageId = "IconPencilTool";
            this.dropDown_writeMemo.ShowImage = true;
            this.dropDown_writeMemo.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_shrink_SelectionChanged);
            // 
            // Ribbon_imageInserter
            // 
            this.Name = "Ribbon_imageInserter";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_imageInserter);
            resources.ApplyResources(this, "$this");
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab_imageInserter.ResumeLayout(false);
            this.tab_imageInserter.PerformLayout();
            this.group_insert.ResumeLayout(false);
            this.group_insert.PerformLayout();
            this.group_setting.ResumeLayout(false);
            this.group_setting.PerformLayout();
            this.group_cell.ResumeLayout(false);
            this.group_cell.PerformLayout();
            this.group_memo.ResumeLayout(false);
            this.group_memo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_imageInserter;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_setting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_insertFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_shrink;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_memo;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_maxSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_maxW;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_maxH;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_direction;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_cell;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton_insert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_insertFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_cell;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_memo;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_insert;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_setW;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_setH;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_setSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_writeCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_writeMemo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_insertLink;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_imageInserter Ribbon1
        {
            get { return this.GetRibbon<Ribbon_imageInserter>(); }
        }
    }
}
