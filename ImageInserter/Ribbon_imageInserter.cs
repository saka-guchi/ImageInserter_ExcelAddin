using System;
using System.Collections.Generic;
using Debug = System.Diagnostics.Debug;
using System.Drawing;
using Image = System.Drawing.Image;
using System.Windows.Forms;
using System.Linq;
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

            // Check params
            if (cells == null)
            {
                return;
            }

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

            // Check params
            if (cells == null)
            {
                return;
            }

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
                    cell = (count == 1) ? cell.Offset[0, 0] : cell.Offset[offsetCol, offsetRow];

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
            Debug.WriteLine("<<< pasteImageOnMemo() >>>");

            // Get UI params
            string write = dropDown_writeMemo.SelectedItem.Tag.ToString();
            string info = (cell.Comment != null) ? cell.Comment.Text() : "";

            // Write information
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
            Debug.WriteLine("Write \"{0}\" to \"{1}\" to Memo", info, write);

            // Read image file to get size
#if true
            System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(imagePath);
            System.Drawing.RotateFlipType rotation = System.Drawing.RotateFlipType.RotateNoneFlipNone;
            foreach( System.Drawing.Imaging.PropertyItem item in bmp.PropertyItems)
            {
                if( item.Id != 0x0112 )
                {
                    continue;
                }
                else
                {
                    switch( item.Value[0] )
                    {
                        case 3:
                            rotation = System.Drawing.RotateFlipType.Rotate180FlipNone;
                            break;
                        case 6:
                            rotation = System.Drawing.RotateFlipType.Rotate90FlipNone;
                            break;
                        case 8:
                            rotation = System.Drawing.RotateFlipType.Rotate270FlipNone;
                            break;
                        default:
                            break;
                    }
                    break;
                }
            }
            float imageW = bmp.Width;
            float imageH = bmp.Height;
            Debug.WriteLine("Image: Path = {0}, (w, h) = ({1:F2}, {2:F2}), rotaion = {3}", imagePath, imageW, imageH, rotation.ToString());

            string tempPath = "";
            if ( rotation != System.Drawing.RotateFlipType.RotateNoneFlipNone)
            {
                System.Drawing.Bitmap tempBmp = (System.Drawing.Bitmap)bmp.Clone();
                tempBmp.RotateFlip(rotation);
                tempPath = System.IO.Path.GetTempFileName();
                tempBmp.Save(tempPath);
                imageW = tempBmp.Width;
                imageH = tempBmp.Height;
                imagePath = tempPath;
                tempBmp.Dispose();
                Debug.WriteLine("Image rotated: Path = {0}, (w, h) = ({1:F2}, {2:F2})", tempPath, imageW, imageH);
            }
            bmp.Dispose();
#else
            System.IO.FileStream fs = System.IO.File.OpenRead(imagePath);
            Image image = Image.FromStream(fs, false, false);
            float imageW = image.Width;
            float imageH = image.Height;
            image.RotateFlip(System.Drawing.RotateFlipType.Rotate90FlipNone);
            image.Dispose();
            Debug.WriteLine("Image: Path = {0}, (w, h) = ({1:F2}, {2:F2})", imagePath, imageW, imageH);
#endif

            // Reduce to the specified maximum size (Keep aspect ratio)
            Debug.WriteLine("Specified max size: (w, h) = ({0:D}, {1:D})", maxW, maxH);
            float shapeW = (maxW != 0) ? maxW : imageW;
            float shapeH = (maxH != 0) ? maxH : imageH;
            float ratioW = imageW / shapeW;
            float ratioH = imageH / shapeH;
            Debug.WriteLine("Resize ratio of Shape: (w, h) = ({0:F2}, {1:F2})", ratioW, ratioH);
            if (ratioW < ratioH)
            {
                shapeW = imageW / ratioH;
            }
            else
            {
                shapeH = imageH / ratioW;
            }
            Debug.WriteLine("Shape: (w, h) = ({0:F2}, {1:F2})", shapeW, shapeH);

            // Initialize Memo
            cell.ClearComments();

            // Add information and image to Memo
            cell.AddComment(info);
            cell.Comment.Shape.Fill.UserPicture(imagePath);
            cell.Comment.Shape.Width = shapeW;
            cell.Comment.Shape.Height = shapeH;

            if (tempPath != "")
            {
                System.IO.File.Delete(imagePath);
            }
        }

        private void pasteImageOnCell(Excel.Worksheet sheet, Excel.Range cell, string imagePath, bool isSetW, bool isSetH)
        {
            Debug.WriteLine("<<< pasteImageOnCell() >>>");

            // Get UI params
            string write = dropDown_writeCell.SelectedItem.Tag.ToString();
            string info = cell.Value;

            // Write information
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
            Debug.WriteLine("Write \"{0}\" to \"{1}\" to Cell", info, write);

            // Calculation ratio for unit conversion
            //  - ColumnWidth: 1 character width (DPI dependent)
            //  - Width: point
            //  - RowHeight: point
            //  - Height: point
            double convRatioW = cell.ColumnWidth / cell.Width;
            Debug.WriteLine("Conversion ratio: ColumnWidth / Width = {0:F2} / {1:F2} = {2:F2}", (double)cell.ColumnWidth, (double)cell.Width, convRatioW);

            // Resize to the specified Cell size
            Debug.WriteLine("Change cell size to specified size: Before");
            Debug.WriteLine(" - Cell: (cw,rh) = ({0:F2},{1:F2})", (double)cell.ColumnWidth, (double)cell.RowHeight);
            if (isSetW)
            {
                cell.ColumnWidth = int.Parse(editBox_setW.Text) * convRatioW;       // point to 1 character width
            }
            if (isSetH)
            {
                cell.RowHeight = int.Parse(editBox_setH.Text);
            }
            Debug.WriteLine("Change cell size to specified size: After");
            Debug.WriteLine(" - Cell: (cw,rh) = ({0:F2},{1:F2})", (double)cell.ColumnWidth, (double)cell.RowHeight);

            // Paste image
            Debug.WriteLine("Paste image to Shape:");
            float cellLeft = (float)cell.Left;
            float cellTop = (float)cell.Top;
            Debug.WriteLine(" - Cell: (Left, Top) = ({0:F2},{1:F2})", cellLeft, cellTop);

            float imageWidth = (float)-1;                                       // Set the width and height to -1 to get the original size
            float imageHeight = (float)-1;
            Debug.WriteLine(" - Image: (Left, Top, Width, Height) = ({0:F2},{1:F2})", imageWidth, imageHeight);

            Excel.Shape shape = sheet.Shapes.AddPicture2(
                imagePath,
                MsoTriState.msoFalse,                                           // LinkToFile: 図を作成元のファイルにリンクするかどうか
                MsoTriState.msoTrue,                                            // SaveWithDocument: リンクされた図が、挿入先のドキュメントと共に保存されるかどうか
                cellLeft, cellTop, imageWidth, imageHeight,
                MsoPictureCompress.msoPictureCompressTrue   // Compress: 画像を挿入するときに圧縮するかどうか
                );
            shape.Left = cellLeft;
            shape.Top = cellTop;
            Application.DoEvents();

            // Shape setting
            shape.LockAspectRatio = MsoTriState.msoFalse;           // When resizing: 高さと幅を個別に変更できる
            shape.Placement = Excel.XlPlacement.xlFreeFloating;     // When deleting or moving cells: 移動とリサイズを行わない

            // Resize Shape
            Debug.WriteLine("Resize shape: ");
            string shrink = dropDown_shrink.SelectedItem.Tag.ToString();    // Image placement settings
            Debug.WriteLine(" - mode: " + shrink);

            // Calculate considering rotation
            float cellRotWidth = (float)cell.Width;
            float cellRotHeight = (float)cell.Height;
            if( ( shape.Rotation.Equals(90f)  ) || (shape.Rotation.Equals(270f)) )
            {
                cellRotWidth = (float)cell.Height;
                cellRotHeight = (float)cell.Width;
            }
            Debug.WriteLine(" - Cell (Rotation): (Width, Height) = ({0:F2},{1:F2})", cellRotWidth, cellRotHeight);

            // Keep aspect and scale
            double resizeRatioW = (double)shape.Width / (double)cellRotWidth;
            double resizeRatioH = (double)shape.Height / (double)cellRotHeight;
            Debug.WriteLine(" - resizeRatioW: shape.Width / cellRotWidth = {0:F2} / {1:F2} = {2:F2}", (double)shape.Width, (double)cellRotWidth, resizeRatioW);
            Debug.WriteLine(" - resizeRatioH: shape.Height / cellRotHeight = {0:F2} / {1:F2} = {2:F2}", (double)shape.Height, (double)cellRotHeight, resizeRatioH);

            Debug.WriteLine("<Before>");
            Debug.WriteLine(" - Shape: (Left, Top) = ({0:F2}, {1:F2})", (float)shape.Left, (float)shape.Top);
            Debug.WriteLine(" - Shape: (w, h) = ({0:F2}, {1:F2})", (double)shape.Width, (double)shape.Height);
            Debug.WriteLine(" - Cell: (cw, rh) = ({0:F2}, {1:F2})", (double)cell.ColumnWidth, (double)cell.RowHeight);

            if (shrink == "fit")
            {
                if(resizeRatioW > resizeRatioH)
                {
                    shape.Width = (float)cellRotWidth;
                    shape.Height /= (float)resizeRatioW;
                }
                else
                {
                    shape.Width /= (float)resizeRatioH;
                    shape.Height = (float)cellRotHeight;
                }
            }
            else if (shrink == "fitW")
            {
                // 幅:セル合わせ、高さ:幅縮小後の画像合わせ
                shape.Width = (float)cellRotWidth;
                shape.Height /= (float)resizeRatioW;
                cell.RowHeight = (float)shape.Height;
            }
            else if (shrink == "fitH")
            {
                // 幅:高さ縮小後の画像合わせ、高さ:セル合わせ
                shape.Width /= (float)resizeRatioH;
                shape.Height = (float)cellRotHeight;
                cell.ColumnWidth = (float)shape.Width * (float)convRatioW;
            }

            if ((shape.Rotation.Equals(90f)) || (shape.Rotation.Equals(270f)))
            {
                float leftMargin = (shape.Width - shape.Height) / 2;
                float topMargin = (shape.Height - shape.Width) / 2;
                shape.Left = cellLeft - leftMargin;
                shape.Top = cellTop - topMargin;
            }

            Debug.WriteLine("<After>");
            Debug.WriteLine(" - Shape: (Left, Top) = ({0:F2},{1:F2})", (float)shape.Left, (float)shape.Top);
            Debug.WriteLine(" - Shape: (w,h) = ({0:F2}, {1:F2})", (double)shape.Width, (double)shape.Height);
            Debug.WriteLine(" - Cell: (cw,rh) = ({0:F2}, {1:F2})", (double)cell.ColumnWidth, (double)cell.RowHeight);
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
            if(Globals.ThisAddIn.Application.Selection == null)
            {
                return null;
            }
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
