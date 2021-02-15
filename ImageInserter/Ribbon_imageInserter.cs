using System;
using System.Collections.Generic;
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
        private void print(
            string message,
            [System.Runtime.CompilerServices.CallerMemberName] string function = "",
            [System.Runtime.CompilerServices.CallerFilePath] string path = "",
            [System.Runtime.CompilerServices.CallerLineNumber] int line = 0
            )
        {
            System.Diagnostics.Trace.WriteLine($"{System.IO.Path.GetFileName(path),15} (L:{line,4}) | {function,-20} | {message}");
        }
        private void Ribbon_imageInserter_Load(object sender, RibbonUIEventArgs e)
        {
            // Load config & Change UI params
            loadSetting();
        }

        private void loadSetting()
        {
            // w/ System.Configuration
            System.Configuration.Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal);
            print($"Load config file: {config.FilePath}");

            Properties.Settings setting = Properties.Settings.Default;

            // When updating, the settings of the previous version are not inherited
            // Inherit the settings of the previous version only for the first time
            if (!Properties.Settings.Default.IsUpgrated)
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.IsUpgrated = true;
            }

            try
            {
                checkBox_cell.Checked = setting.checkBox_cell;
                checkBox_memo.Checked = setting.checkBox_memo;
                checkBox_setSize.Checked = setting.checkBox_setSize;
                editBox_setW.Text = $"{setting.editBox_setW}";
                editBox_setH.Text = $"{setting.editBox_setH}";
                dropDown_shrink.SelectedItemIndex = setting.dropDown_shrink;
                dropDown_writeCell.SelectedItemIndex = setting.dropDown_writeCell;
                dropDown_deleteCell.SelectedItemIndex = setting.dropDown_deleteCell;
                dropDown_direction.SelectedItemIndex = setting.dropDown_direction;
                checkBox_cell.Checked = setting.checkBox_maxSize;
                editBox_maxW.Text = $"{setting.editBox_maxW:D}";
                editBox_maxH.Text = $"{setting.editBox_maxH:D}";
                dropDown_writeMemo.SelectedItemIndex = setting.dropDown_writeMemo;
                dropDown_deleteMemo.SelectedItemIndex = setting.dropDown_deleteMemo;
            }
            catch (Exception ex)
            {
                print($"<<< ERROR >>>: {ex}");
            }

            switch (setting.splitButton_insert)
            {
                case "select":
                    changeEvent_splitButton(splitButton_insert, button_insertFile);
                    splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_insertFile_Click);
                    break;
                case "link":
                    changeEvent_splitButton(splitButton_insert, button_insertLink);
                    splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_insertLink_Click);
                    break;
                case "folder":
                    changeEvent_splitButton(splitButton_insert, button_insertFolder);
                    splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_insertFolder_Click);
                    break;
                default:
                    break;
            }
            switch (setting.splitButton_delete)
            {
                case "select":
                    changeEvent_splitButton(splitButton_delete, button_deleteSelection);
                    splitButton_delete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_deleteSelection_Click);
                    break;
                case "all":
                    changeEvent_splitButton(splitButton_delete, button_deleteAll);
                    splitButton_delete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_deleteAll_Click);
                    break;
                default:
                    break;
            }
            updateControlState(true);
        }

        private void saveSetting()
        {
            Properties.Settings setting = Properties.Settings.Default;

            setting.checkBox_cell = checkBox_cell.Checked;
            setting.checkBox_memo = checkBox_memo.Checked;
            setting.checkBox_setSize = checkBox_setSize.Checked;
            setting.editBox_setW = int.Parse(editBox_setW.Text);
            setting.editBox_setH = int.Parse(editBox_setH.Text);
            setting.dropDown_shrink = dropDown_shrink.SelectedItemIndex;
            setting.dropDown_writeCell = dropDown_writeCell.SelectedItemIndex;
            setting.dropDown_deleteCell = dropDown_deleteCell.SelectedItemIndex;
            setting.dropDown_direction = dropDown_direction.SelectedItemIndex;
            setting.checkBox_maxSize = checkBox_cell.Checked;
            setting.editBox_maxW = int.Parse(editBox_maxW.Text);
            setting.editBox_maxH = int.Parse(editBox_maxH.Text);
            setting.dropDown_writeMemo = dropDown_writeMemo.SelectedItemIndex;
            setting.dropDown_deleteMemo = dropDown_deleteMemo.SelectedItemIndex;
            setting.splitButton_insert = splitButton_insert.Tag.ToString();
            setting.splitButton_delete = splitButton_delete.Tag.ToString();

            setting.Save();

            // w/ System.Configuration
            System.Configuration.Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal);
            print($"Save config file: {config.FilePath}");
        }

        private void changeEvent_splitButton(RibbonSplitButton btnDst, RibbonButton btnSrc)
        {
            try
            {
                // Delete all because registration event is unknown
                btnDst.Click -= button_insertFile_Click;
                btnDst.Click -= button_insertFolder_Click;
                btnDst.Click -= button_insertLink_Click;
                btnDst.Click -= button_deleteSelection_Click;
                btnDst.Click -= button_deleteAll_Click;

                btnDst.Label = btnSrc.Label;
                btnDst.Tag = btnSrc.Tag;
                btnDst.OfficeImageId = btnSrc.OfficeImageId;
            }
            catch (Exception ex)
            {
                print($"<<< ERROR >>>: {ex}");
            }
        }

        private void button_insertFile_Click(object sender, RibbonControlEventArgs e)
        {
            // Change UI params
            changeEvent_splitButton(splitButton_insert, button_insertFile);
            splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_insertFile_Click);

            // Get UI params
            Excel.Worksheet sheet = null;
            Excel.Range cell = null;
            string imagePath = "";

            try
            {
                sheet = getActiveSheet();
                cell = getActiveCell();
                imagePath = getImagePathFromDialog();
            }
            catch (Exception ex)
            {
                print($"ERROR: {ex}");
                return;
            }

            // Check params
            if (imagePath == null)
            {
                return;
            }

            // Disable UI
            switchControlState(false);

            pasteImage(sheet, cell, imagePath);

            // Enable UI
            switchControlState(true);
        }

        private void button_insertLink_Click(object sender, RibbonControlEventArgs e)
        {
            // Change UI params
            changeEvent_splitButton(splitButton_insert, button_insertLink);
            splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_insertLink_Click);

            // Get UI params
            Excel.Worksheet sheet = null;
            Excel.Range cellsSelect = null;
            Excel.Range cellsFill = null;
            Excel.Range cells = null;

            try
            {
                sheet = getActiveSheet();
                cellsSelect = getSelection();
                cellsFill = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Excel.XlSpecialCellsValue.xlTextValues);     // Constants and Text
                cells = sheet.Application.Intersect(cellsSelect, cellsFill);
            }
            catch (Exception ex)
            {
                print($"ERROR: {ex}");
                return;
            }

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
            // Change UI params
            changeEvent_splitButton(splitButton_insert, button_insertFolder);
            splitButton_insert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_insertFolder_Click);

            // Get UI params
            Excel.Worksheet sheet = null;
            Excel.Range cell = null;
            string direction = "";
            string folderPath = "";

            try
            {
                sheet = getActiveSheet();
                cell = getActiveCell();
                direction = dropDown_direction.SelectedItem.Tag.ToString();
                folderPath = getFolderPath(cell);
            }
            catch (Exception ex)
            {
                print($"ERROR: {ex}");
                return;
            }

            // Check params
            if (folderPath == "")
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
            // Change UI params
            changeEvent_splitButton(splitButton_delete, button_deleteSelection);
            splitButton_delete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_deleteSelection_Click);

            // Get UI params
            Excel.Worksheet sheet = null;
            Excel.Range cells = null;
            bool checkCell = false;
            bool checkMemo = false;
            bool checkCellKeep = false;
            bool checkMemoKeep = false;

            try
            {
                sheet = getActiveSheet();
                cells = getSelection();
                checkCell = checkBox_cell.Checked;
                checkMemo = checkBox_memo.Checked;
                checkCellKeep = (dropDown_deleteCell.SelectedItem.Tag.ToString() == "keep");
                checkMemoKeep = (dropDown_deleteMemo.SelectedItem.Tag.ToString() == "keep");
            }
            catch (Exception ex)
            {
                print($"ERROR: {ex}");
                return;
            }

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
            // Change UI params
            changeEvent_splitButton(splitButton_delete, button_deleteAll);
            splitButton_delete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(button_deleteAll_Click);

            // Get UI params
            Excel.Application app = null;
            Excel.Worksheet sheet = null;
            Excel.Range cells = null;
            Excel.Range cellsComments = null;
            Excel.Range cellsConstants = null;
            bool checkCell = false;
            bool checkMemo = false;
            bool checkCellKeep = false;
            bool checkMemoKeep = false;
            try
            {
                app = getApplication();
                sheet = getActiveSheet();
                try
                {
                    cellsComments = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeComments);
                }
                catch (Exception)
                {
                }
                try
                {
                    cellsConstants = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants);
                }
                catch (Exception)
                {
                }
                if (cellsComments != null)
                {
                    if (cellsConstants != null)
                    {
                        cells = app.Union(cellsComments, cellsConstants);
                    }
                    else
                    {
                        cells = cellsComments;
                    }
                }
                else if (cellsConstants != null)
                {
                    cells = cellsConstants;
                }
                checkCell = checkBox_cell.Checked;
                checkMemo = checkBox_memo.Checked;
                checkCellKeep = (dropDown_deleteCell.SelectedItem.Tag.ToString() == "keep");
                checkMemoKeep = (dropDown_deleteMemo.SelectedItem.Tag.ToString() == "keep");
            }
            catch (Exception ex)
            {
                print($"ERROR: {ex}");
                return;
            }

            // Check params
            if (cells == null)
            {
                return;
            }

            // Disable UI
            switchControlState(false);

            // Delete all images
            deleteAllImages(sheet, cells, checkCell, checkMemo, checkCellKeep, checkMemoKeep);
        }

        private void switchUIBehavior(bool enable)
        {
            try
            {
                Globals.ThisAddIn.Application.Interactive = enable;
                Globals.ThisAddIn.Application.ScreenUpdating = enable;
            }
            catch (Exception ex)
            {
                print($"<<< ERROR >>>: {ex}");
            }
        }

        private void updateControlState(bool enable)
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                print($"<<< ERROR >>>: {ex}");
            }
        }

        private void switchControlState(bool enable)
        {
            if (enable)
            {
                switchUIBehavior(enable);
            }

            updateControlState(enable);

            if (!enable)
            {
                switchUIBehavior(enable);
            }
        }

        private void checkBox_setSize_Click(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
            switchControlState(true);
        }
        private void checkBox_maxSize_Click(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
            switchControlState(true);
        }

        private void dropDown_shrink_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
            switchControlState(true);
        }

        private async void pasteLinkedImages(Excel.Worksheet sheet, Excel.Range cells)
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

                // Bring the window to the front using your own process. w/ Microsoft.VisualBasic.dll
                System.Diagnostics.Process p = System.Diagnostics.Process.GetCurrentProcess();
                Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
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
            waitDlg.Count = $"{count:G}/{countMax:G}";
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

                // Bring the window to the front using your own process. w/ Microsoft.VisualBasic.dll
                System.Diagnostics.Process p = System.Diagnostics.Process.GetCurrentProcess();
                Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
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

                // Bring the window to the front using your own process. w/ Microsoft.VisualBasic.dll
                System.Diagnostics.Process p = System.Diagnostics.Process.GetCurrentProcess();
                Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
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
                    if (isAll)
                    {
                        isValueClear = true;
                    }
                    else
                    {
                        // Check the intersection of Shape and Selected cells
                        if (sheet.Application.Intersect(shapeRange, selectedCells) != null)
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

                // Bring the window to the front using your own process. w/ Microsoft.VisualBasic.dll
                System.Diagnostics.Process p = System.Diagnostics.Process.GetCurrentProcess();
                Microsoft.VisualBasic.Interaction.AppActivate(p.Id);
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
            print($"Image: Path = {imagePath}, (w, h) = ({imageW:F2}, {imageH:F2})");
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
                print($"Image: rotaion = {rotation}");
            }
            catch (Exception ex)
            {
                print($"ERROR: {ex}");
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
                    print($"Rotated image: Path = {tempPath}, (w, h) = ({imageW:F2}, {imageH:F2})");
                }
                catch (Exception ex)
                {
                    print($"ERROR tempBmp: {ex}");
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
            catch (Exception ex)
            {
                print($"<<< ERROR >>>: {ex}");
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
                    print($"Add hyperlink to Cell: \"{imageOrgPath}\"");
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
            else if ((type == "name") || (type == "nameLink"))
            {
                info = System.IO.Path.GetFileName(imagePath);
            }
            else if ((type == "path") || (type == "pathLink"))
            {
                info = imagePath;
            }
            return info;
        }

        private void pasteImageOnCell(Excel.Worksheet sheet, Excel.Range cell, string writeInfo, string imagePath, float imageW, float imageH, bool isSetW, bool isSetH)
        {
            cell.Value = writeInfo;
            print($"Write \"{writeInfo}\" to Cell");

            // Calculation ratio for unit conversion
            //  - ColumnWidth: 1 character width (DPI dependent)
            //  - Width: point
            //  - RowHeight: point
            //  - Height: point
            double convRatioW = cell.ColumnWidth / cell.Width;
            print($"Conversion ratio: ColumnWidth / Width = {(float)cell.ColumnWidth:F2} / {(float)cell.Width:F2} = {convRatioW:F2}");

            // Resize to the specified Cell size
            print("Change cell size to specified size: Before");
            print($" - Cell: (cw,rh) = ({(float)cell.ColumnWidth:F2},{(float)cell.RowHeight:F2})");
            if (isSetW)
            {
                cell.ColumnWidth = int.Parse(editBox_setW.Text) * convRatioW;       // point to 1 character width
            }
            if (isSetH)
            {
                cell.RowHeight = int.Parse(editBox_setH.Text);
            }
            print("Change cell size to specified size: After");
            print($" - Cell: (cw,rh) = ({(float)cell.ColumnWidth:F2},{(float)cell.RowHeight:F2})");

            // Paste image
            print("Paste image to Shape:");
            float cellLeft = (float)cell.Left;
            float cellTop = (float)cell.Top;
            print($" - Cell: (Left, Top) = ({cellLeft:F2},{cellTop:F2})");
            print($" - Image: (w, h) = ({imageW:F2},{imageH:F2})");

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
                shape.Placement = Excel.XlPlacement.xlMove;                // Shape is moved with the cells.

                // Resize Shape
                print("Resize shape: ");
                string shrink = dropDown_shrink.SelectedItem.Tag.ToString();    // Image placement settings
                print($" - mode: {shrink}");

                // Calculate considering rotation
                float cellRotWidth = (float)cell.Width;
                float cellRotHeight = (float)cell.Height;
                if ((shape.Rotation.Equals(90f)) || (shape.Rotation.Equals(270f)))
                {
                    cellRotWidth = (float)cell.Height;
                    cellRotHeight = (float)cell.Width;
                }
                print($" - Cell (Rotation): (Width, Height) = ({cellRotWidth:F2},{cellRotHeight:F2})");

                // Keep aspect and scale
                double resizeRatioW = (double)shape.Width / (double)cellRotWidth;
                double resizeRatioH = (double)shape.Height / (double)cellRotHeight;
                print($" - resizeRatioW: shape.Width / cellRotWidth = {shape.Width:F2} / {cellRotWidth:F2} = {resizeRatioW:F2}");
                print($" - resizeRatioH: shape.Height / cellRotHeight = {shape.Height:F2} / {cellRotHeight:F2} = {resizeRatioH:F2}");

                print("<Before>");
                print($" - Shape: (Left, Top) = ({shape.Left:F2}, {shape.Top:F2})");
                print($" - Shape: (w, h) = ({shape.Width:F2}, {shape.Height:F2})");
                print($" - Cell: (cw, rh) = ({(float)cell.ColumnWidth:F2}, {(float)cell.RowHeight:F2})");

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
            catch (Exception ex)
            {
                print($"ERROR shape: {ex}");
            }
            print("<After>");
            print($" - Shape: (Left, Top) = ({shape.Left:F2},{shape.Top:F2})");
            print($" - Shape: (w,h) = ({shape.Width:F2}, {shape.Height:F2})");
            print($" - Cell: (cw,rh) = ({(float)cell.ColumnWidth:F2}, {(float)cell.RowHeight:F2})");
        }

        private void pasteImageOnMemo(Excel.Worksheet sheet, Excel.Range cell, string writeInfo, string imagePath, float imageW, float imageH, int maxW, int maxH)
        {
            // Reduce to the specified maximum size (Keep aspect ratio)
            print($"Specified max size: (w, h) = ({maxW:D}, {maxH:D})");
            float shapeW = (maxW != 0) ? maxW : imageW;
            float shapeH = (maxH != 0) ? maxH : imageH;
            float ratioW = imageW / shapeW;
            float ratioH = imageH / shapeH;
            print($"Resize ratio of Shape: (w, h) = ({ratioW:F2}, {ratioH:F2})");

            if (ratioW < ratioH)
            {
                shapeW = imageW / ratioH;
            }
            else
            {
                shapeH = imageH / ratioW;
            }
            print($"Shape: (w, h) = ({shapeW:F2}, {shapeH:F2})");

            // Initialize Memo
            cell.ClearComments();

            // Add information and image to Memo
            print($"Write \"{writeInfo}\" to Memo");
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
            string[] extCheck = { ".jpg", ".jpeg", ".bmp", ".png", ".gif" };
            if (extCheck.Contains(ext, System.StringComparer.OrdinalIgnoreCase))
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
            catch (Exception ex)
            {
                print($"ERROR tempBmp: {ex}");
            }
            ofd.Dispose();

            return imagePath;
        }

        private Excel.Application getApplication()
        {
            if (Globals.ThisAddIn.Application == null)
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
            if (Globals.ThisAddIn.Application.ActiveSheet == null)
            {
                return null;
            }
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            return activeSheet;
        }

        private Excel.Range getSelection()
        {
            if (Globals.ThisAddIn.Application.Selection == null)
            {
                return null;
            }
            Excel.Range selection = Globals.ThisAddIn.Application.Selection;
            return selection;
        }

        private Excel.Range getActiveCell()
        {
            if (Globals.ThisAddIn.Application.ActiveCell == null)
            {
                return null;
            }
            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            return activeCell;
        }

        private string getFolderPath(Excel.Range cell)
        {
            string folderPath = "";
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
            string folderPath = "";
            FolderSelectDialog dlg = new FolderSelectDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                folderPath = dlg.Path;
            }
            return folderPath;
        }

        private void checkBox_cell_Click(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
        }

        private void checkBox_memo_Click(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
        }

        private void editBox_setW_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // Check params
            try
            {
                int w = int.Parse(editBox_setW.Text);
            }
            catch (Exception ex)
            {
                print($"<<< Parse error >>>: {ex}");
                editBox_setW.Text = "15";
                return;
            }
            saveSetting();
        }

        private void editBox_setH_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // Check params
            try
            {
                int h = int.Parse(editBox_setH.Text);
            }
            catch (Exception ex)
            {
                print($"<<< Parse error >>>: {ex}");
                editBox_setH.Text = "15";
                return;
            }
            saveSetting();
        }

        private void dropDown_writeCell_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
        }

        private void dropDown_deleteCell_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
        }

        private void dropDown_direction_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
        }

        private void editBox_maxW_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // Check params
            try
            {
                int h = int.Parse(editBox_maxW.Text);
            }
            catch (Exception ex)
            {
                print($"<<< Parse error >>>: {ex}");
                editBox_maxW.Text = "512";
                return;
            }
            saveSetting();
        }

        private void editBox_maxH_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // Check params
            try
            {
                int h = int.Parse(editBox_maxH.Text);
            }
            catch (Exception ex)
            {
                print($"<<< Parse error >>>: {ex}");
                editBox_maxH.Text = "512";
                return;
            }
            saveSetting();
        }

        private void dropDown_writeMemo_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
        }

        private void dropDown_deleteMemo_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            saveSetting();
        }
    }
}
