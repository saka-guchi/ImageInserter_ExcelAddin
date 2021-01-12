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
            switchControlState(false);

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
                return;
            }

            int offsetCol = (direction == "under") ? 1 : 0;
            int offsetRow = (direction == "right") ? 1 : 0;

            // Get files in specified folder
            string[] exts = { ".jpg", ".jpeg", ".bmp", ".png", ".gif" };
            List<string> imgList = GetFilesInFolder(folderPath, exts);

            // Disable UI
            switchControlState(false);

            // Paste images on cells
            pasteImages(sheet, cell, imgList, offsetCol, offsetRow);
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
            switchControlState(false);

            // Delete Images in selection
            deleteImagesInSelection(sheet, cells, checkCell, checkMemo, checkCellKeep, checkMemoKeep);
        }

        private void button_deleteAll_Click(object sender, RibbonControlEventArgs e)
        {
            // Get UI params
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cells = (sheet.Cells.Comment == null) ? sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeComments): null;
            bool checkCell = checkBox_cell.Checked;
            bool checkMemo = checkBox_memo.Checked;
            bool checkCellKeep = (dropDown_deleteCell.SelectedItem.Tag.ToString() == "keep");
            bool checkMemoKeep = (dropDown_deleteMemo.SelectedItem.Tag.ToString() == "keep");

            // Disable UI
            switchControlState(false);

            // Delete all images
            deleteAllImages(sheet, cells, checkCell, checkMemo, checkCellKeep, checkMemoKeep);
        }

        private void switchControlState(bool enable)
        {
            if (enable)
            {
                Globals.ThisAddIn.Application.Interactive = enable;
                Globals.ThisAddIn.Application.ScreenUpdating = enable;
                Globals.ThisAddIn.Application.ActiveSheet.Application.ScreenUpdating = enable;
                Application.DoEvents();
            }

            // Get all UI controls
            foreach (RibbonGroup group in Globals.Ribbons.Ribbon1.tab_imageInserter.Groups)
            {
                foreach (RibbonControl ctrl in group.Items)
                {
                    ctrl.Enabled = enable;
                }
            }

            // Correspond individually
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

            if (!enable)
            {
                Globals.ThisAddIn.Application.ActiveSheet.Application.ScreenUpdating = enable;
                Globals.ThisAddIn.Application.ScreenUpdating = enable;
                Globals.ThisAddIn.Application.Interactive = enable;
                Application.DoEvents();
            }
        }

        private void checkBox_setSize_Click(object sender, RibbonControlEventArgs e)
        {
            switchControlState(true);
        }
        private void checkBox_maxSize_Click(object sender, RibbonControlEventArgs e)
        {
            switchControlState(true);
        }

        private void dropDown_shrink_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            switchControlState(true);
        }

        private async void pasteLinkedImages(Excel.Worksheet sheet, Excel. Range cells)
        {
            int countMax = cells.Count;
            await System.Threading.Tasks.Task.Run(() =>
            {
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
                    waitDialogDisplay(waitDlg, count, countMax);

                    // Processing
                    string imagePath = getImagePathFromCell(cell);
                    if (imagePath == null)
                    {
                        imagePath = getImagePathFromHyperlink(cell);
                    }

                    // Check params
                    if (imagePath != null)
                    {
                        pasteImage(sheet, cell, imagePath);
                    }

                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();

                // Enable UI
                switchControlState(true);
            }
            );
        }

        private List<string> GetFilesInFolder(string path, string[] exts)
        {
            List<string> files = System.IO.Directory.GetFiles(path)
                .Where(f => exts.Contains(System.IO.Path.GetExtension(f), System.StringComparer.OrdinalIgnoreCase))
                .ToList();
            return files;
        }

        private void waitDialogDisplay(WaitDialog waitDlg, int count, int countMax)
        {
            // Display progress message
            waitDlg.Count = String.Format("{0}/{1}", count.ToString(), countMax.ToString());
            waitDlg.Percentage = String.Format("{0:P}", (float)count / (float)countMax);
            waitDlg.PerformStep();
            Application.DoEvents();
        }

        private async void pasteImages(Excel.Worksheet sheet, Excel.Range cell, List<string> imgList, int offsetCol, int offsetRow)
        {
            int countMax = imgList.Count;
            await System.Threading.Tasks.Task.Run(() =>
            {
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
                    waitDialogDisplay(waitDlg, count, countMax);

                    // Processing
                    cell = (count == 1) ? cell.Offset[0, 0] : cell.Offset[offsetCol, offsetRow];

                    // Check params
                    if (imagePath != null)
                    {
                        pasteImage(sheet, cell, imagePath);
                    }

                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();

                // Enable UI
                switchControlState(true);
            }
            );
        }

        private async void deleteImagesInSelection(Excel.Worksheet sheet, Excel.Range cells, bool checkCell, bool checkMemo, bool checkCellKeep, bool checkMemoKeep)
        {
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
                    waitDialogDisplay(waitDlg, count, countMax);

                    // Processing
                    deleteImage(sheet, shape, cells, checkCell, checkMemo, checkCellKeep, checkMemoKeep, false);

                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();

                // Enable UI
                switchControlState(true);
            }
            );
        }

        private void deleteImage(Excel.Worksheet sheet, Excel.Shape shape, Excel.Range selectedCells, bool checkCell, bool checkMemo, bool checkCellKeep, bool checkMemoKeep, bool isAll)
        {
            // Target: Cell
            if (checkCell)
            {
                if (shape.Type == MsoShapeType.msoPicture)     // image
                {
                    // Get Cell under Shape
                    Excel.Range shapeRange = shape.TopLeftCell;

                    // Erase the contents of the cell
                    bool isValueClear = false;
                    if ( isAll )
                    {
                        isValueClear = true;
                    }
                    else
                    {
                        // Check the intersection of Shape and Selected cells
                        if(sheet.Application.Intersect(shapeRange, selectedCells) != null)
                        {
                            isValueClear = true;
                        }
                    }
                    if (isValueClear)
                    {
                        if (!checkCellKeep)
                        {
                            shapeRange.Value = "";
                            shapeRange.Hyperlinks.Delete();
                        }
                        // Delete an image in a cell
                        shape.Delete();

                        // Skip checkMemo
                        return;
                    }
                }
            }

            // Target: Memo
            if (checkMemo)
            {
                if (shape.Type == MsoShapeType.msoComment)  // memo
                {
                    // Get cells containing comment
                    foreach (Excel.Range cell in selectedCells)
                    {
                        if (cell.Comment == null)
                        {
                            continue;
                        }

                        // Compare ID
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
        }

        private async void deleteAllImages(Excel.Worksheet sheet, Excel.Range cells, bool checkCell, bool checkMemo, bool checkCellKeep, bool checkMemoKeep)
        {
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
                    waitDialogDisplay(waitDlg, count, countMax);

                    // Processing
                    deleteImage(sheet, shape, cells, checkCell, checkMemo, checkCellKeep, checkMemoKeep, true);

                    count++;
                }

                // Progress bar: Close
                waitDlg.Close();

                // Enable UI
                switchControlState(true);
            }
            );

        }

        private string preloadImage(ref string imagePath, ref float imageW, ref float imageH)
        {
            // Read image file to get size and rotation
            System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(imagePath);
            System.Drawing.RotateFlipType rotation = System.Drawing.RotateFlipType.RotateNoneFlipNone;
            imageW = bmp.Width;
            imageH = bmp.Height;
            Debug.WriteLine("Image: Path = {0}, (w, h) = ({1:F2}, {2:F2})", imagePath, imageW, imageH);
            try
            {
                foreach (System.Drawing.Imaging.PropertyItem item in bmp.PropertyItems)
                {
                    if (item.Id != 0x0112)
                    {
                        continue;
                    }
                    else
                    {
                        switch (item.Value[0])
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
                Debug.WriteLine("Image: rotaion = {0}", rotation.ToString());
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR: {0}", e);
            }

            // Comment.Shape.Fil.UserPicture() is displayed without rotation.
            // Create a rotated image and display it instead
            string tempPath = "";
            if (rotation != System.Drawing.RotateFlipType.RotateNoneFlipNone)
            {
                System.Drawing.Bitmap tempBmp = null;
                try
                {
                    tempBmp = (System.Drawing.Bitmap)bmp.Clone();
                    tempBmp.RotateFlip(rotation);
                    tempPath = System.IO.Path.GetTempFileName();
                    tempBmp.Save(tempPath);
                    imageW = tempBmp.Width;
                    imageH = tempBmp.Height;
                    imagePath = tempPath;
                    Debug.WriteLine("Image rotated: Path = {0}, (w, h) = ({1:F2}, {2:F2})", tempPath, imageW, imageH);
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR tempBmp: {0}", e);
                }
                tempBmp.Dispose();
            }
            bmp.Dispose();

            return tempPath;
        }

        private void pasteImage(Excel.Worksheet sheet, Excel.Range cell, string imageOrgPath)
        {
            // Get UI params
            bool pasteCell = false;
            bool pasteMemo = false;
            bool setMaxSize = false;
            int maxW = 0;
            int maxH = 0;
            bool isSetW = false;
            bool isSetH = false;
            string writeInfoCell = "";
            bool isLink = false;
            string writeInfoMemo = "";
            string comment = "";
            try
            {
                pasteCell = checkBox_cell.Checked;
                pasteMemo = checkBox_memo.Checked;
                setMaxSize = checkBox_maxSize.Checked;
                if (setMaxSize)
                {
                    maxW = int.Parse(editBox_maxW.Text);
                    maxH = int.Parse(editBox_maxH.Text);
                }
                isSetW = editBox_setW.Enabled;
                isSetH = editBox_setH.Enabled;
                writeInfoCell = dropDown_writeCell.SelectedItem.Tag.ToString();
                isLink = ((writeInfoCell == "nameLink") || (writeInfoCell == "pathLink")) ? true : false;
                writeInfoMemo = dropDown_writeMemo.SelectedItem.Tag.ToString();
                comment = (cell.Comment != null) ? cell.Comment.Text() : "";
            }
            catch(Exception e)
            {
                Debug.WriteLine("<<< ERROR >>> Get UI params: {0}", e);
                return;
            }

            // imageOriginalPath: Path to the original picture
            // imagePath: Path to the load picture (original or temporary)
            // tempPath: Path to the rotated picture
            float imageW = 0f;
            float imageH = 0f;
            string imagePath = imageOrgPath;
            string tempPath = preloadImage(ref imagePath, ref imageW, ref imageH);

            // Paste image on Cell
            if (pasteCell)
            {
                string writeInfo = getWriteInfo(imageOrgPath, writeInfoCell, cell.Value);
                pasteImageOnCell(sheet, cell, writeInfo, imagePath, imageW, imageH, isSetW, isSetH);
                if (isLink)
                {
                    sheet.Hyperlinks.Add(cell, imageOrgPath);
                    Debug.WriteLine("Add hyperlink to Cell: \"{0}\"", imageOrgPath);
                }
            }

            // Paste image on Memo
            if (pasteMemo)
            {
                string writeInfo = getWriteInfo(imageOrgPath, writeInfoMemo, comment);
                pasteImageOnMemo(sheet, cell, writeInfo, imagePath, imageW, imageH, maxW, maxH);
            }

            // Delete temporary file.
            if (tempPath != "")
            {
                System.IO.File.Delete(tempPath);
            }
        }

        private string getWriteInfo(string imagePath, string type, string infoInit)
        {
            string info = infoInit;
            if (type == "none")
            {
                ;
            }
            else if ( (type == "name") || (type == "nameLink") )
            {
                info = System.IO.Path.GetFileName(imagePath);
            }
            else if ( (type == "path") || (type == "pathLink") )
            {
                info = imagePath;
            }
            return info;
        }

        private void pasteImageOnCell(Excel.Worksheet sheet, Excel.Range cell, string writeInfo, string imagePath, float imageW, float imageH, bool isSetW, bool isSetH)
        {
            Debug.WriteLine("<<< pasteImageOnCell() >>>");

            cell.Value = writeInfo;
            Debug.WriteLine("Write \"{0}\" to Cell", writeInfo);

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
            Debug.WriteLine(" - Image: (w, h) = ({0:F2},{1:F2})", imageW, imageH);

            Excel.Shape shape = sheet.Shapes.AddPicture2(
                imagePath,
                MsoTriState.msoFalse,                                           // LinkToFile: Picture will be linked to the file from which it was created
                MsoTriState.msoTrue,                                            // SaveWithDocument: Linked picture will be saved with the document
                cellLeft, cellTop, imageW, imageH,
                MsoPictureCompress.msoPictureCompressTrue   // Compress: Picture should be compressed when inserted
                );
            try
            {
                shape.Left = cellLeft;
                shape.Top = cellTop;
                Application.DoEvents();

                // Shape setting
                shape.LockAspectRatio = MsoTriState.msoFalse;           // Shape retains its original proportions
                shape.Placement = Excel.XlPlacement.xlFreeFloating;     // Shape is free floating

                // Resize Shape
                Debug.WriteLine("Resize shape: ");
                string shrink = dropDown_shrink.SelectedItem.Tag.ToString();    // Image placement settings
                Debug.WriteLine(" - mode: " + shrink);

                // Calculate considering rotation
                float cellRotWidth = (float)cell.Width;
                float cellRotHeight = (float)cell.Height;
                if ((shape.Rotation.Equals(90f)) || (shape.Rotation.Equals(270f)))
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
                    if (resizeRatioW > resizeRatioH)
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
                    // Width: Cell width, Height: Resized Image height
                    shape.Width = (float)cellRotWidth;
                    shape.Height /= (float)resizeRatioW;
                    cell.RowHeight = (float)shape.Height;
                }
                else if (shrink == "fitH")
                {
                    // Width: Resized Cell width, Height: Cell height
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
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR shape: {0}", e);
            }
            Debug.WriteLine("<After>");
            Debug.WriteLine(" - Shape: (Left, Top) = ({0:F2},{1:F2})", (float)shape.Left, (float)shape.Top);
            Debug.WriteLine(" - Shape: (w,h) = ({0:F2}, {1:F2})", (double)shape.Width, (double)shape.Height);
            Debug.WriteLine(" - Cell: (cw,rh) = ({0:F2}, {1:F2})", (double)cell.ColumnWidth, (double)cell.RowHeight);
        }

        private void pasteImageOnMemo(Excel.Worksheet sheet, Excel.Range cell, string writeInfo, string imagePath, float imageW, float imageH, int maxW, int maxH)
        {
            Debug.WriteLine("<<< pasteImageOnMemo() >>>");

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
            Debug.WriteLine("Write \"{0}\" to Memo", writeInfo);
            cell.AddComment(writeInfo);
            cell.Comment.Shape.Fill.UserPicture(imagePath);
            cell.Comment.Shape.Width = shapeW;
            cell.Comment.Shape.Height = shapeH;
        }

        private string getImagePathFromCell(Excel.Range cell)
        {
            string imagePath = null;
            if (checkImagePath(cell.Text))
            {
                imagePath = cell.Text;
            }
            return imagePath;
        }

        private string getImagePathFromHyperlink(Excel.Range cell)
        {
            string imagePath = null;
            foreach (Excel.Hyperlink link in cell.Hyperlinks)
            {
                if (checkImagePath(link.Address))
                {
                    imagePath = link.Address;
                    break;
                }
            }
            return imagePath;
        }

        private bool checkImagePath(string path)
        {
            // Checking the existence and extension
            bool isExist = false;
            string ext = System.IO.Path.GetExtension(path);
            string[] extCheck = {".jpg",".jpeg", ".bmp", ".png", ".gif"};
            if ( extCheck.Contains(ext, System.StringComparer.OrdinalIgnoreCase))
            {
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

            // Create an instance of the OpenFileDialog class
            OpenFileDialog ofd = new OpenFileDialog();
            try
            {
                ofd.Filter = "Image file - 画像ファイル(*.jpeg;*.jpg;*.bmp;*.png;*.gif)|*.jpeg;*.jpg;*.bmp;*.png;*.gif|All files - すべてのファイル(*.*)|*.*";
                ofd.FilterIndex = 1;                    // 1: Image file, 2: All files
                ofd.Title = "Please select a file - ファイルを選択してください";
                ofd.RestoreDirectory = true;     // Restore to the previously selected directory
                ofd.ShowHelp = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    imagePath = ofd.FileName;
                }
            }
            catch(Exception e)
            {
                Debug.WriteLine("ERROR tempBmp: {0}", e);
            }
            ofd.Dispose();

            return imagePath;
        }

        private Excel.Application getApplication()
        {
            if(Globals.ThisAddIn.Application == null)
            {
                return null;
            }
            Excel.Application application = Globals.ThisAddIn.Application;
            return application;
        }

        private Excel.Workbook getActiveWorkBook()
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
            {
                return null;
            }
            Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            return activeWorkbook;
        }

        private Excel.Worksheet getActiveSheet()
        {
            if(Globals.ThisAddIn.Application.ActiveSheet == null)
            {
                return null;
            }
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            return activeSheet;
        }

        private Excel.Range getSelection()
        {
            if(Globals.ThisAddIn.Application.Selection == null)
            {
                return null;
            }
            Excel.Range selection = Globals.ThisAddIn.Application.Selection;
            return selection;
        }

        private Excel.Range getActiveCell()
        {
            if(Globals.ThisAddIn.Application.ActiveCell == null)
            {
                return null;
            }
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            return activeCell;
        }

        private string getFolderPath(Excel.Range cell)
        {
            string folderPath = null;
            if (isExistFolderPath(cell.Text))
            {
                folderPath = cell.Text;
            }
            else
            {
                folderPath = getFolderPathFromDialog();
            }
            return folderPath;
        }

        private bool isExistFolderPath(string path)
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
        }
    }
}
