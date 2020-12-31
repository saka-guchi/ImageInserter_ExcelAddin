
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            this.tab_imageInserter = this.Factory.CreateRibbonTab();
            this.group_insert = this.Factory.CreateRibbonGroup();
            this.group_setting = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.checkBox_cell = this.Factory.CreateRibbonCheckBox();
            this.checkBox_memo = this.Factory.CreateRibbonCheckBox();
            this.group_cell = this.Factory.CreateRibbonGroup();
            this.checkBox_setSize = this.Factory.CreateRibbonCheckBox();
            this.editBox_setW = this.Factory.CreateRibbonEditBox();
            this.editBox_setH = this.Factory.CreateRibbonEditBox();
            this.dropDown_shrink = this.Factory.CreateRibbonDropDown();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.dropDown_direction = this.Factory.CreateRibbonDropDown();
            this.group_memo = this.Factory.CreateRibbonGroup();
            this.checkBox_maxSize = this.Factory.CreateRibbonCheckBox();
            this.editBox_maxW = this.Factory.CreateRibbonEditBox();
            this.editBox_maxH = this.Factory.CreateRibbonEditBox();
            this.splitButton_insert = this.Factory.CreateRibbonSplitButton();
            this.button_insertFile = this.Factory.CreateRibbonButton();
            this.button_insertFolder = this.Factory.CreateRibbonButton();
            this.dropDown_writeCell = this.Factory.CreateRibbonDropDown();
            this.dropDown_writeMemo = this.Factory.CreateRibbonDropDown();
            this.button_insertLink = this.Factory.CreateRibbonButton();
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
            this.tab_imageInserter.Label = "画像挿入";
            this.tab_imageInserter.Name = "tab_imageInserter";
            // 
            // group_insert
            // 
            this.group_insert.Items.Add(this.splitButton_insert);
            this.group_insert.Label = "画像の挿入";
            this.group_insert.Name = "group_insert";
            // 
            // group_setting
            // 
            this.group_setting.Items.Add(this.label1);
            this.group_setting.Items.Add(this.checkBox_cell);
            this.group_setting.Items.Add(this.checkBox_memo);
            this.group_setting.Label = "全般設定";
            this.group_setting.Name = "group_setting";
            // 
            // label1
            // 
            this.label1.Label = "挿入する場所";
            this.label1.Name = "label1";
            // 
            // checkBox_cell
            // 
            this.checkBox_cell.Checked = true;
            this.checkBox_cell.Label = "セル";
            this.checkBox_cell.Name = "checkBox_cell";
            // 
            // checkBox_memo
            // 
            this.checkBox_memo.Checked = true;
            this.checkBox_memo.Label = "メモ";
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
            this.group_cell.Label = "セル設定";
            this.group_cell.Name = "group_cell";
            // 
            // checkBox_setSize
            // 
            this.checkBox_setSize.Label = "セルサイズ指定";
            this.checkBox_setSize.Name = "checkBox_setSize";
            this.checkBox_setSize.SuperTip = "画像挿入時にセルのサイズを変更します";
            this.checkBox_setSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_setSize_Click);
            // 
            // editBox_setW
            // 
            this.editBox_setW.Enabled = false;
            this.editBox_setW.Label = "　幅";
            this.editBox_setW.MaxLength = 4;
            this.editBox_setW.Name = "editBox_setW";
            this.editBox_setW.Text = "15";
            // 
            // editBox_setH
            // 
            this.editBox_setH.Enabled = false;
            this.editBox_setH.Label = "高さ";
            this.editBox_setH.MaxLength = 4;
            this.editBox_setH.Name = "editBox_setH";
            this.editBox_setH.Text = "15";
            // 
            // dropDown_shrink
            // 
            ribbonDropDownItemImpl1.Label = "セル内に収める";
            ribbonDropDownItemImpl1.OfficeImageId = "BackgroundImageGallery";
            ribbonDropDownItemImpl1.Tag = "fit";
            ribbonDropDownItemImpl2.Label = "セル幅に合わせる";
            ribbonDropDownItemImpl2.OfficeImageId = "CellHeight";
            ribbonDropDownItemImpl2.Tag = "fitW";
            ribbonDropDownItemImpl3.Label = "セル高さに合わせる";
            ribbonDropDownItemImpl3.OfficeImageId = "GroupTableCellFormat";
            ribbonDropDownItemImpl3.Tag = "fitH";
            this.dropDown_shrink.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown_shrink.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown_shrink.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown_shrink.Label = "縮小方法";
            this.dropDown_shrink.Name = "dropDown_shrink";
            this.dropDown_shrink.OfficeImageId = "DiagramResizeClassic";
            this.dropDown_shrink.ShowImage = true;
            this.dropDown_shrink.SuperTip = "セルの中に画像をどのように挿入するかを指定します";
            this.dropDown_shrink.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_shrink_SelectionChanged);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // label2
            // 
            this.label2.Label = "複数セルの配置";
            this.label2.Name = "label2";
            // 
            // dropDown_direction
            // 
            ribbonDropDownItemImpl7.Label = "下";
            ribbonDropDownItemImpl7.OfficeImageId = "ChartNavDrillDown";
            ribbonDropDownItemImpl7.Tag = "bottom";
            ribbonDropDownItemImpl8.Label = "右";
            ribbonDropDownItemImpl8.OfficeImageId = "OrgChartReportMoveRight";
            ribbonDropDownItemImpl8.Tag = "right";
            this.dropDown_direction.Items.Add(ribbonDropDownItemImpl7);
            this.dropDown_direction.Items.Add(ribbonDropDownItemImpl8);
            this.dropDown_direction.Label = "配置方向";
            this.dropDown_direction.Name = "dropDown_direction";
            this.dropDown_direction.SuperTip = "複数の画像を挿入する際に選択したセルからどの方向に画像を追加していくかを指定します";
            // 
            // group_memo
            // 
            this.group_memo.Items.Add(this.checkBox_maxSize);
            this.group_memo.Items.Add(this.editBox_maxW);
            this.group_memo.Items.Add(this.editBox_maxH);
            this.group_memo.Items.Add(this.dropDown_writeMemo);
            this.group_memo.Label = "メモ設定";
            this.group_memo.Name = "group_memo";
            // 
            // checkBox_maxSize
            // 
            this.checkBox_maxSize.Checked = true;
            this.checkBox_maxSize.Label = "最大サイズ";
            this.checkBox_maxSize.Name = "checkBox_maxSize";
            this.checkBox_maxSize.SuperTip = "メモに挿入する画像のサイズの上限を指定します";
            this.checkBox_maxSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_maxSize_Click);
            // 
            // editBox_maxW
            // 
            this.editBox_maxW.Label = "　幅";
            this.editBox_maxW.MaxLength = 4;
            this.editBox_maxW.Name = "editBox_maxW";
            this.editBox_maxW.Text = "512";
            // 
            // editBox_maxH
            // 
            this.editBox_maxH.Label = "高さ";
            this.editBox_maxH.MaxLength = 4;
            this.editBox_maxH.Name = "editBox_maxH";
            this.editBox_maxH.Text = "512";
            // 
            // splitButton_insert
            // 
            this.splitButton_insert.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton_insert.Items.Add(this.button_insertFile);
            this.splitButton_insert.Items.Add(this.button_insertLink);
            this.splitButton_insert.Items.Add(this.button_insertFolder);
            this.splitButton_insert.Label = "指定した画像を挿入";
            this.splitButton_insert.Name = "splitButton_insert";
            this.splitButton_insert.OfficeImageId = "RestoreImageSize";
            this.splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertFile_Click);
            // 
            // button_insertFile
            // 
            this.button_insertFile.Label = "指定した画像を挿入";
            this.button_insertFile.Name = "button_insertFile";
            this.button_insertFile.OfficeImageId = "RestoreImageSize";
            this.button_insertFile.ShowImage = true;
            this.button_insertFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertFile_Click);
            // 
            // button_insertFolder
            // 
            this.button_insertFolder.Label = "フォルダ内の画像を挿入";
            this.button_insertFolder.Name = "button_insertFolder";
            this.button_insertFolder.OfficeImageId = "ApplyImageBackgroundTile";
            this.button_insertFolder.ShowImage = true;
            this.button_insertFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertFolder_Click);
            // 
            // dropDown_writeCell
            // 
            ribbonDropDownItemImpl4.Label = "しない";
            ribbonDropDownItemImpl4.OfficeImageId = "CancelRequest";
            ribbonDropDownItemImpl4.Tag = "none";
            ribbonDropDownItemImpl5.Label = "ファイル名";
            ribbonDropDownItemImpl5.OfficeImageId = "FileNew";
            ribbonDropDownItemImpl5.Tag = "file";
            ribbonDropDownItemImpl6.Label = "パス";
            ribbonDropDownItemImpl6.OfficeImageId = "GroupImapFolderOptions";
            ribbonDropDownItemImpl6.Tag = "path";
            this.dropDown_writeCell.Items.Add(ribbonDropDownItemImpl4);
            this.dropDown_writeCell.Items.Add(ribbonDropDownItemImpl5);
            this.dropDown_writeCell.Items.Add(ribbonDropDownItemImpl6);
            this.dropDown_writeCell.Label = "情報書込";
            this.dropDown_writeCell.Name = "dropDown_writeCell";
            this.dropDown_writeCell.OfficeImageId = "IconPencilTool";
            this.dropDown_writeCell.ShowImage = true;
            this.dropDown_writeCell.SuperTip = "セルにファイル名やパスを書き込みます";
            this.dropDown_writeCell.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_shrink_SelectionChanged);
            // 
            // dropDown_writeMemo
            // 
            ribbonDropDownItemImpl9.Label = "しない";
            ribbonDropDownItemImpl9.OfficeImageId = "CancelRequest";
            ribbonDropDownItemImpl9.Tag = "none";
            ribbonDropDownItemImpl10.Label = "ファイル名";
            ribbonDropDownItemImpl10.OfficeImageId = "FileNew";
            ribbonDropDownItemImpl10.Tag = "file";
            ribbonDropDownItemImpl11.Label = "パス";
            ribbonDropDownItemImpl11.OfficeImageId = "GroupImapFolderOptions";
            ribbonDropDownItemImpl11.Tag = "path";
            this.dropDown_writeMemo.Items.Add(ribbonDropDownItemImpl9);
            this.dropDown_writeMemo.Items.Add(ribbonDropDownItemImpl10);
            this.dropDown_writeMemo.Items.Add(ribbonDropDownItemImpl11);
            this.dropDown_writeMemo.Label = "情報書込";
            this.dropDown_writeMemo.Name = "dropDown_writeMemo";
            this.dropDown_writeMemo.OfficeImageId = "IconPencilTool";
            this.dropDown_writeMemo.ShowImage = true;
            this.dropDown_writeMemo.SuperTip = "メモにファイル名やパスを書き込みます";
            this.dropDown_writeMemo.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown_shrink_SelectionChanged);
            // 
            // button_insertLink
            // 
            this.button_insertLink.Label = "セルのリンク先画像を挿入";
            this.button_insertLink.Name = "button_insertLink";
            this.button_insertLink.OfficeImageId = "OmsImageFromClip";
            this.button_insertLink.ShowImage = true;
            this.button_insertLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_insertLink_Click);
            // 
            // Ribbon_imageInserter
            // 
            this.Name = "Ribbon_imageInserter";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_imageInserter);
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
