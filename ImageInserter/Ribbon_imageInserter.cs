using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using Debug = System.Diagnostics.Debug;
using Image = System.Drawing.Image;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using FolderSelectDialogForm;
using WaitDialogForm;

namespace ImageInserter
{
    public partial class Ribbon_imageInserter
    {
        private void Ribbon_imageInserter_Load(object sender, RibbonUIEventArgs e)
        {
            // Initialize UI
            dropDown_writeCell.SelectedItemIndex = 2;       // Write path in cell when images inserting
            dropDown_writeMemo.SelectedItemIndex = 1;   // Write file name in memo when images inserting
            dropDown_deleteCell.SelectedItemIndex = 1;     // Keep cell contents when deleting
            dropDown_deleteMemo.SelectedItemIndex = 0; // Keep memo contents when deleting
        }

        private void button_insertFile_Click(object sender, RibbonControlEventArgs e)
        {
            // Get UI params
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cell = getActiveCell();
            string imagePath = getImagePathFromDialog();

            // Check params
            if (imagePath == null)
            {
                MessageBox.Show(
                    "Image file does not exist\n画像ファイルが存在しません",
                    "ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                    );
                return;
            }

            pasteImage(sheet, cell, imagePath);
        }

        private void button_insertLink_Click(object sender, RibbonControlEventArgs e)
        {
            // Get UI params
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cells = getSelection();

            // Disable UI
            switch_control_state(false);

            // Paste Linked images on cells
            pasteLinkedImages(sheet, cells);
        }

        private void button_insertFolder_Click(object sender, RibbonControlEventArgs e)
        {
            // Get UI params
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cell = getActiveCell();
            string direction = dropDown_direction.SelectedItem.Tag.ToString();
            string folderPath = getFolderPath(cell);

            // Check params
            if (folderPath == null)
            {
                MessageBox.Show(
                    "Folder does not exist\nフォルダが存在しません",
                    "ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                    );
                return;
            }
            int offsetCol = (direction == "under") ? 1 : 0;
            int offsetRow = (direction == "right") ? 1 : 0;

            // Get files in specified folder
            string[] exts = { ".jpg", ".jpeg" };
            List<string> imgList = GetFilesInFolder(folderPath, exts);

            // Disable UI
            switch_control_state(false);

            // Paste images on cells
            pasteImages(imgList, sheet, cell, offsetCol, offsetRow);
        }

        private void button_deleteSelection_Click(object sender, RibbonControlEventArgs e)
        {
            // Get UI params
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cells = getSelection();
            bool checkCell = checkBox_cell.Checked;
            bool checkMemo = checkBox_memo.Checked;
            bool checkCellKeep = (dropDown_deleteCell.SelectedItem.Tag.ToString() == "keep");
            bool checkMemoKeep = (dropDown_deleteMemo.SelectedItem.Tag.ToString() == "keep");

            // Disable UI
            switch_control_state(false);

            // Delete Images in selection
            deleteImagesInSelection(sheet, cells, checkCell, checkMemo, checkCellKeep, checkMemoKeep);
        }

        private void button_deleteAll_Click(object sender, RibbonControlEventArgs e)
        {
            // Get UI params
            Excel.Worksheet sheet = getActiveSheet();
            bool checkCell = checkBox_cell.Checked;
            bool checkMemo = checkBox_memo.Checked;
            bool checkCellKeep = (dropDown_deleteCell.SelectedItem.Tag.ToString() == "keep");
            bool checkMemoKeep = (dropDown_deleteMemo.SelectedItem.Tag.ToString() == "keep");

            // Disable UI
            switch_control_state(false);

            // Delete all images
            deleteAllImages(sheet, checkCell, checkMemo, checkCellKeep, checkMemoKeep);
        }

        private void switch_control_state(bool enable)
        {
            foreach(RibbonGroup group in Globals.Ribbons.Ribbon1.tab_imageInserter.Groups)
            {
                foreach (RibbonControl ctrl in group.Items)
                {
                    ctrl.Enabled = enable;
                }
            }

            if (enable)
            {
                bool max_w = false;
                bool max_h = false;
                if (checkBox_maxSize.Checked)
                {
                    max_w = true;
                    max_h = true;
                }
                editBox_maxW.Enabled = max_w;
                editBox_maxH.Enabled = max_h;

                bool set_w = false;
                bool set_h = false;
                if (checkBox_setSize.Checked)
                {
                    string shrink = dropDown_shrink.SelectedItem.Tag.ToString();
                    if (shrink == "fit")
                    {
                        set_w = true;
                        set_h = true;
                    }
                    else if (shrink == "fitW")
                    {
                        set_w = true;
                    }
                    else if (shrink == "fitH")
                    {
                        set_h = true;
                    }
                }
                editBox_setW.Enabled = set_w;
                editBox_setH.Enabled = set_h;
            }
        }

        private void checkBox_setSize_Click(object sender, RibbonControlEventArgs e)
        {
            switch_control_state(true);
        }
        private void checkBox_maxSize_Click(object sender, RibbonControlEventArgs e)
        {
            switch_control_state(true);
        }

        private void dropDown_shrink_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            switch_control_state(true);
        }

        private async void pasteLinkedImages(Excel.Worksheet sheet, Excel. Range cells)
        {
            // Stop screen updating
            Excel.Application app = getApplication();
            app.ScreenUpdating = false;

            int countMax = cells.Count;
            await System.Threading.Tasks.Task.Run(() => {
                // Progress bar: Setting
                WaitDialog waitDlg = new WaitDialog();
                waitDlg.ProgressMax = countMax;

                // Progress bar: Show
                waitDlg.Show();
                Application.DoEvents();

                int count = 1;
                foreach (Excel.Range cell in cells)
                {
                    // Stop decision
                    if (waitDlg.IsAborting == true)
                    {
                        break;
                    }

                    // Display progress message
                    waitDlg.Count = String.Format("{0}/{1}", count.ToString(), countMax.ToString());
                    waitDlg.Percentage = String.Format("{0:P}", (float)count / (float)countMax);
                    waitDlg.PerformStep();
                    Application.DoEvents();

                    // Processing
                    string imagePath = cell.Text;

                    // Paste image
                    if ( checkImagePath(imagePath) == true )
                    {
                        pasteImage(sheet, cell, imagePath);
                    }
                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();
                Application.DoEvents();
            });

            // Statr screen updating
            app.ScreenUpdating = true;

            // Enable UI
            switch_control_state(true);
        }

        private List<string> GetFilesInFolder(string path, string[] exts)
        {
            string[] extCheck = { ".jpg", ".jpeg" };
            List<string> files = System.IO.Directory.GetFiles(path)
                .Where(f => extCheck.Contains(System.IO.Path.GetExtension(f), System.StringComparer.OrdinalIgnoreCase))
                .ToList();
            return files;
        }

        private async void pasteImages(List<string> imgList, Excel.Worksheet sheet, Excel.Range cell, int offsetCol, int offsetRow)
        {
            // Stop screen updating
            Excel.Application app = getApplication();
            app.ScreenUpdating = false;

            int countMax = imgList.Count;
            await System.Threading.Tasks.Task.Run(() => {
                // Progress bar: Setting
                WaitDialog waitDlg = new WaitDialog();
                waitDlg.ProgressMax = countMax;

                // Progress bar: Show
                waitDlg.Show();
                Application.DoEvents();

                int count = 1;
                foreach (string imagePath in imgList)
                {
                    // Stop decision
                    if (waitDlg.IsAborting == true)
                    {
                        break;
                    }

                    // Display progress message
                    waitDlg.Count = String.Format("{0}/{1}", count.ToString(), countMax.ToString());
                    waitDlg.Percentage = String.Format("{0:P}", (float)count / (float)countMax);
                    waitDlg.PerformStep();
                    Application.DoEvents();

                    // Processing
                    cell = (count == 0) ? cell.Offset[0, 0] : cell.Offset[offsetCol, offsetRow];

                    // Paste image
                    if (checkImagePath(imagePath) == true)
                    {
                        pasteImage(sheet, cell, imagePath);
                    }
                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();
                Application.DoEvents();
            });

            // Statr screen updating
            app.ScreenUpdating = true;

            // Enable UI
            switch_control_state(true);
        }

        private async void deleteImagesInSelection(Excel.Worksheet sheet, Excel.Range cells, bool checkCell, bool checkMemo, bool checkCellKeep, bool checkMemoKeep)
        {
            // Stop screen updating
            Excel.Application app = getApplication();
            app.ScreenUpdating = false;

            int countMax = sheet.Shapes.Count;
            await System.Threading.Tasks.Task.Run(() =>
            {
                // Progress bar: Setting
                WaitDialog waitDlg = new WaitDialog();
                waitDlg.ProgressMax = countMax;

                // Progress bar: Show
                waitDlg.Show();
                Application.DoEvents();

                int count = 1;
                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    // Stop decision
                    if (waitDlg.IsAborting == true)
                    {
                        break;
                    }

                    // Display progress message
                    waitDlg.Count = String.Format("{0}/{1}", count.ToString(), countMax.ToString());
                    waitDlg.Percentage = String.Format("{0:P}", (float)count / (float)countMax);
                    waitDlg.PerformStep();
                    Application.DoEvents();

                    // Processing
                    if (checkCell)
                    {
                        // Target: Cell
                        if (shape.Type == MsoShapeType.msoLinkedPicture)
                        {
                            // Intersect: 2つ以上の範囲の長方形の交差を表すRangeオブジェクトを返す
                            Excel.Range shapeRange = shape.TopLeftCell;
                            if (sheet.Application.Intersect(shapeRange, cells) != null)
                            {
                                if (!checkCellKeep)
                                {
                                    shapeRange.Value = "";
                                }
                                shape.Delete();          // Delete an image in a cell
                                count++;
                                continue;                   // Skip checkMemo
                            }
                        }
                    }
                    if (checkMemo)
                    {
                        // Target: Memo
                        if (shape.Type == MsoShapeType.msoComment)
                        {
                            // 選択範囲内のコメントを含むセルのコメントのShapeのIDと比較
                            foreach (Excel.Range cell in cells)
                            {
                                if (cell.Comment == null)
                                {
                                    continue;
                                }
                                if (cell.Comment.Shape.ID == shape.ID)
                                {
                                    if (checkMemoKeep)
                                    {
                                        shape.Fill.Solid();     // Delete image in memo
                                    }
                                    else
                                    {
                                        cell.Comment.Delete();  // Delete memo
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();
                Application.DoEvents();
            });

            // Statr screen updating
            app.ScreenUpdating = true;

            // Enable UI
            switch_control_state(true);
        }

        private async void deleteAllImages(Excel.Worksheet sheet, bool checkCell, bool checkMemo, bool checkCellKeep, bool checkMemoKeep)
        {
            // Stop screen updating
            Excel.Application app = getApplication();
            app.ScreenUpdating = false;

            int countMax = sheet.Shapes.Count;
            await System.Threading.Tasks.Task.Run(() => {
                // Progress bar: Setting
                WaitDialog waitDlg = new WaitDialog();
                waitDlg.ProgressMax = countMax;

                // Progress bar: Show
                waitDlg.Show();
                Application.DoEvents();

                int count = 1;
                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    // Stop decision
                    if (waitDlg.IsAborting == true)
                    {
                        break;
                    }

                    // Display progress message
                    waitDlg.Count = String.Format("{0}/{1}", count.ToString(), countMax.ToString());
                    waitDlg.Percentage = String.Format("{0:P}", (float)count / (float)countMax);
                    waitDlg.PerformStep();
                    Application.DoEvents();

                    // Processing
                    if (checkCell)
                    {
                        // Target: Cell
                        if (shape.Type == MsoShapeType.msoLinkedPicture)
                        {
                            if (!checkCellKeep)
                            {
                                Excel.Range shapeRange = shape.TopLeftCell;
                                shapeRange.Value = "";
                            }
                            shape.Delete();          // Delete an image in a cell
                            count++;
                            continue;                   // Skip checkMemo
                        }
                    }
                    if (checkMemo)
                    {
                        // Target: Memo
                        if (shape.Type == MsoShapeType.msoComment)
                        {
                            if (checkMemoKeep)
                            {
                                shape.Fill.Solid();     // Delete image in memo
                            }
                            else
                            {
                                // すべてのコメントを含むセルのShapeのIDと比較
                                foreach (Excel.Range cell in sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeComments))
                                {
                                    if (cell.Comment.Shape.ID == shape.ID)
                                    {
                                        cell.Comment.Delete();  // Delete memo
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();
                Application.DoEvents();
            });

            // Statr screen updating
            app.ScreenUpdating = true;

            // Enable UI
            switch_control_state(true);
        }

        private void pasteImage(Excel.Worksheet sheet, Excel.Range cell, string imagePath)
        {
            if (checkBox_cell.Checked)
            {
                // セルに画像貼付
                pasteImageOnCell(sheet, cell, imagePath, editBox_setW.Enabled, editBox_setH.Enabled);
            }
            if (checkBox_memo.Checked)
            {
                // 最大サイズの取得
                int maxW = 0;
                int maxH = 0;
                if (checkBox_maxSize.Checked)
                {
                    maxW = int.Parse(editBox_maxW.Text);
                    maxH = int.Parse(editBox_maxH.Text);
                }

                // メモに画像貼付
                pasteImageOnMemo(sheet, cell, imagePath, maxW, maxH);
            }
        }

        private void pasteImageOnMemo(Excel.Worksheet sheet, Excel.Range cell, string imagePath, int maxW, int maxH)
        {
            System.IO.FileStream fs = System.IO.File.OpenRead(imagePath);
            Image image = Image.FromStream(fs, false, false);
            float w = image.Width;
            float h = image.Height;
            image.Dispose();

            // アスペクト比を維持して最大サイズまで縮小
            float limW = (maxW != 0) ? maxW : w;
            float limH = (maxH != 0) ? maxH : h;
            double ratioW = (double)w / limW;
            double ratioH = (double)h / limH;
            if (ratioW > ratioH)
            {
                w = (float)limW;
                h /= (float)ratioW;
            }
            else
            {
                w /= (float)ratioH;
                h = (float)limH;
            }

            // 情報書込
            string write = dropDown_writeMemo.SelectedItem.Tag.ToString();
            string info = (cell.Comment != null)? cell.Comment.Text(): "";
            if (write == "none")
            {
                ;
            }
            else if (write == "name")
            {
                info = System.IO.Path.GetFileName(imagePath);
            }
            else if (write == "path")
            {
                info = imagePath;
            }

            cell.ClearComments();
            cell.AddComment(info);
            cell.Comment.Shape.Fill.UserPicture(imagePath);
            cell.Comment.Shape.Width = w;
            cell.Comment.Shape.Height = h;
        }

        private void pasteImageOnCell(Excel.Worksheet sheet, Excel.Range cell, string imagePath, bool isSetW, bool isSetH)
        {
            // 情報書込
            string write = dropDown_writeCell.SelectedItem.Tag.ToString();
            string info = cell.Value;
            if (write == "none")
            {
                ;
            }
            else if (write == "name")
            {
                info = System.IO.Path.GetFileName(imagePath);
            }
            else if (write == "path")
            {
                info = imagePath;
            }
            cell.Value = info;

            // セルのサイズ変更
            double ratioC = cell.ColumnWidth / cell.Width;
            double ratioR = cell.RowHeight / cell.Height;
            Debug.WriteLine("ratioC = \t" + ratioC);
            Debug.WriteLine("ratioR = \t" + ratioR);
            if (isSetW)
            {
                cell.ColumnWidth = int.Parse(editBox_setW.Text) * ratioC;
            }
            if (isSetH)
            {
                cell.RowHeight = int.Parse(editBox_setH.Text) * ratioR;
            }

            float left = (float)cell.Left;
            float top = (float)cell.Top;

            // 読み込む画像の幅と高さが不明な場合、ゼロにして後で等倍に拡大
            float width = 0.0f;
            float height = 0.0f;

            Excel.Shape shape = sheet.Shapes.AddPicture2(
                imagePath,
                MsoTriState.msoTrue,                        // LinkToFile: 図を作成元のファイルにリンクするかどうか
                MsoTriState.msoTrue,                        // SaveWithDocument: リンクされた図が、挿入先のドキュメントと共に保存されるかどうか
                left, top, width, height,
                MsoPictureCompress.msoPictureCompressTrue   // Compress: 画像を挿入するときに圧縮するかどうか
                );

            // 縦横比を固定
//            shape.LockAspectRatio = MsoTriState.msoTrue;
            shape.LockAspectRatio = MsoTriState.msoFalse;

            // セルに合わせて移動・サイズ変更
//            shape.Placement = Excel.XlPlacement.xlMoveAndSize;
            shape.Placement = Excel.XlPlacement.xlFreeFloating;

            // 元のサイズに戻す
            shape.ScaleHeight(1.0f, MsoTriState.msoTrue);
            shape.ScaleWidth(1.0f, MsoTriState.msoTrue);

            // 配置設定を取得
            string shrink = dropDown_shrink.SelectedItem.Tag.ToString();

            // アスペクトを維持して拡大縮小する
            double ratioW = (double)shape.Width / cell.Width;
            double ratioH = (double)shape.Height / cell.Height;
            Debug.WriteLine("ratioW = \t" + ratioW);
            Debug.WriteLine("ratioH = \t" + ratioH);

            Debug.WriteLine("BEFORE");
            Debug.WriteLine("shape.Width = \t" + shape.Width);
            Debug.WriteLine("shape.Height = \t" + shape.Height);
            Debug.WriteLine("cell.Width = \t" + (float)cell.Width);
            Debug.WriteLine("cell.Height = \t" + (float)cell.Height);
            Debug.WriteLine("cell.ColumnWidth = \t" + (float)cell.ColumnWidth);
            Debug.WriteLine("cell.RowHeight = \t" + (float)cell.RowHeight);

            if (shrink == "fit")
            {
                if( ratioW > ratioH )
                {
                    shape.Width = (float)cell.Width;
                    shape.Height /= (float)ratioW;
                }
                else
                {
                    shape.Width /= (float)ratioH;
                    shape.Height = (float)cell.Height;
                }
            }
            else if (shrink == "fitW")
            {
                // 幅:セル合わせ、高さ:幅縮小後の画像合わせ
                shape.Width = (float)cell.Width;
                shape.Height /= (float)ratioW;
                cell.RowHeight = (double)shape.Height * ratioR;
            }
            else if (shrink == "fitH")
            {
                // 幅:高さ縮小後の画像合わせ、高さ:セル合わせ
                shape.Width /= (float)ratioH;
                shape.Height = (float)cell.Height;
                cell.ColumnWidth = (double)shape.Width * ratioC;
            }
            Debug.WriteLine("AFTER");
            Debug.WriteLine("shape.Width = \t" + shape.Width);
            Debug.WriteLine("shape.Height = \t" + shape.Height);
            Debug.WriteLine("cell.Width = \t" + (float)cell.Width);
            Debug.WriteLine("cell.Height = \t" + (float)cell.Height);
            Debug.WriteLine("cell.ColumnWidth = \t" + (float)cell.ColumnWidth);
            Debug.WriteLine("cell.RowHeight = \t" + (float)cell.RowHeight);
        }


        private string getImagePathFromCell(Excel.Range cell)
        {
            string imagePath = null;
            if (checkImagePath(cell.Text))
            {
                // セルから取得
                imagePath = cell.Text;
            }
            return imagePath;
        }

        private bool checkImagePath(string path)
        {
            bool isExist = false;
            string ext = System.IO.Path.GetExtension(path);
            string[] extCheck = {".jpg",".jpeg"};
            if ( extCheck.Contains(ext, System.StringComparer.OrdinalIgnoreCase))
            {
                // 存在確認
                if (System.IO.File.Exists(path))
                {
                    isExist = true;
                }
            }
            return isExist;
        }

        private string getImagePathFromDialog()
        {
            string imagePath = null;

            // OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            // はじめに表示されるフォルダを指定する
            // 指定しない（空の文字列）の時は、現在のディレクトリが表示される
//          ofd.InitialDirectory = @"C:\";

            // [ファイルの種類]に表示される選択肢を指定する
            // 指定しないとすべてのファイルが表示される
            ofd.Filter = "画像ファイル(*.jpegl;*.jpg)|*.jpeg;*.jpg|すべてのファイル(*.*)|*.*";

            // [ファイルの種類]ではじめに選択されるものを指定する
            ofd.FilterIndex = 1;    // 1:画像ファイル, 2:すべてのファイル

            // タイトルを設定する
            ofd.Title = "開くファイルを選択してください";

            // ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;

            // 存在しないファイルの名前が指定されたとき警告を表示する
            // デフォルトでTrueなので指定する必要はない
//          ofd.CheckFileExists = true;

            // 存在しないパスが指定されたとき警告を表示する
            // デフォルトでTrueなので指定する必要はない
//          ofd.CheckPathExists = true;

            // ダイアログ ボックスに [ヘルプ] ボタンを表示する
            ofd.ShowHelp = true;

            // ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                imagePath = ofd.FileName;
            }

            // オブジェクトを破棄する
            ofd.Dispose();

            return imagePath;
        }

        private Excel.Application getApplication()
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            return application;
        }

        // アクティブワークブックの取得
        private Excel.Workbook getActiveWorkBook()
        {
            Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            return activeWorkbook;
        }

        // アクティブシートの取得
        private Excel.Worksheet getActiveSheet()
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            return activeSheet;
        }

        // 選択範囲の取得
        private Excel.Range getSelection()
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection;
            return selection;
        }

        // 選択セルの取得
        private Excel.Range getActiveCell()
        {
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            return activeCell;
        }

        private string getFolderPath(Excel.Range cell)
        {
            string folderPath = null;
            if (checkFolderPath(cell.Text))
            {
                // セルから取得
                folderPath = cell.Text;
            }
            else
            {
                // ダイアログから取得
                folderPath = getFolderPathFromDialog();
            }
            return folderPath;
        }

        // フォルダの存在確認
        private bool checkFolderPath(string path)
        {
            bool isExist = false;
            if (System.IO.Directory.Exists(path))
            {
                isExist = true;
            }
            return isExist;
        }

        private string getFolderPathFromDialog()
        {
            string folderPath = null;

            FolderSelectDialog dlg = new FolderSelectDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                folderPath = dlg.Path;
            }
            return folderPath;

#if false   // default dialog (Tree type)
            // OpenFileDialogクラスのインスタンスを作成
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            // ダイアログの説明文を指定する
            fbd.Description = "画像フォルダを選択してください";

            // はじめに表示されるフォルダを指定する
            // 指定しない（空の文字列）の時は、現在のディレクトリが表示される
            // ofd.SelectedPath = @"C:\";

            // 「新しいフォルダーを作成する」ボタンを表示する
            fbd.ShowNewFolderButton = true;

            // ダイアログを表示する
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                folderPath = fbd.SelectedPath;
            }

            // オブジェクトを破棄する
            fbd.Dispose();

            return folderPath;
#endif
        }
    }
}
