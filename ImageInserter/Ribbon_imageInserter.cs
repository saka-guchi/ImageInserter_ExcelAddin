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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // 初期設定
            dropDown_writeCell.SelectedItemIndex = 2;       // パス
            dropDown_writeMemo.SelectedItemIndex = 1;   // ファイル名
        }

        private void checkBox_setSize_Click(object sender, RibbonControlEventArgs e)
        {
            dropDown_shrink_SelectionChanged(sender, null);
        }
        private void checkBox_maxSize_Click(object sender, RibbonControlEventArgs e)
        {
            editBox_maxW.Enabled = (checkBox_maxSize.Checked) ? true : false;
            editBox_maxH.Enabled = (checkBox_maxSize.Checked) ? true : false;
        }

        private void dropDown_shrink_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (checkBox_setSize.Checked)
            {
                string shrink = dropDown_shrink.SelectedItem.Tag.ToString();
                if (shrink == "fit")
                {
                    editBox_setW.Enabled = true;
                    editBox_setH.Enabled = true;
                }
                else if (shrink == "fitW")
                {
                    editBox_setW.Enabled = true;
                    editBox_setH.Enabled = false;
                }
                else if (shrink == "fitH")
                {
                    editBox_setW.Enabled = false;
                    editBox_setH.Enabled = true;
                }
            }
            else
            {
                editBox_setW.Enabled = false;
                editBox_setH.Enabled = false;
            }
        }

        private void button_insertFile_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cell = getActiveCell();

            string imagePath = getImagePathFromDialog(); 
            Debug.WriteLine(imagePath);

            if (imagePath == null)
            {
                MessageBox.Show(
                    "画像ファイルが存在しません",
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
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cells = getSelection();

            // プログレスバーを表示開始
            splitButton_insert.Enabled = false;

            WaitDialog waitDlg = new WaitDialog();
            waitDlg.ProgressMax = cells.Count;
            waitDlg.Show();
            Application.DoEvents();

            int count = 0;
            foreach (Excel.Range cell in cells)
            {
                string imagePath = getImagePathFromCell(cell);
                Debug.WriteLine(imagePath);

                if (imagePath != null)
                {
                    pasteImage(sheet, cell, imagePath);
                }
#if false
                else
                {
                    MessageBox.Show(
                    "画像ファイルが存在しません",
                        "ERROR",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                        );
                    return;
                }
#endif

                // 処理中止かどうかをチェック
                if (waitDlg.IsAborting == true)
                {
                    break;
                }
                else
                {
                    // 進行状況ダイアログのメーターを設定
                    waitDlg.Msg = cells.Count.ToString() + "件中: " + count.ToString() + "件目";
                    waitDlg.PerformStep();
                    Application.DoEvents();
                }
            }
            Debug.WriteLine("count = \t" + count);

            // 最終メッセージを表示
            if (waitDlg.DialogResult == DialogResult.Abort)
            {
                waitDlg.Msg = "処理を中断しました。";
            }
            else
            {
                waitDlg.Msg = "処理を完了しました。";
            }
            Application.DoEvents();
            System.Threading.Thread.Sleep(100);
            waitDlg.Close();

            // プログレスバーを表示終了
            splitButton_insert.Enabled = true;
        }

        private void button_insertFolder_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = getActiveSheet();
            Excel.Range cell = getActiveCell();

            string folderPath = getFolderPath(cell);
            Debug.WriteLine(folderPath);

            if (folderPath == null)
            {
                MessageBox.Show(
                    "フォルダが存在しません",
                    "ERROR",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                    );
                return;
            }

            // フォルダ内のファイルを取得
//            string[] files = System.IO.Directory.GetFiles(folderPath);
#if true   // 高速だがcountが使えない
            string[] extCheck = { ".jpg", ".jpeg" };
            List<string> files = System.IO.Directory.GetFiles(folderPath)
                .Where(f => extCheck.Contains(System.IO.Path.GetExtension(f), System.StringComparer.OrdinalIgnoreCase))
                .ToList();
#endif

            // ファイルを1つずつ処理
            string direction = dropDown_direction.SelectedItem.Tag.ToString();
            int offsetCol = 0;
            int offsetRow = 0;
            if (direction == "under")
            {
                offsetCol = 1;
            }
            if (direction == "right")
            {
                offsetRow = 1;
            }

            // プログレスバーを表示開始
            splitButton_insert.Enabled = false;

            WaitDialog waitDlg = new WaitDialog();
            waitDlg.ProgressMax = files.Count;
            waitDlg.Show();
            Application.DoEvents();

            int count = 0;
            foreach (string imagePath in files)
            {
//                if (checkImagePath(imagePath))
                {
                    if (count == 0)
                    {
                        cell = cell.Offset[0, 0];
                    }
                    else
                    {
                        cell = cell.Offset[offsetCol, offsetRow];
                    }
                    pasteImage(sheet, cell, imagePath);
                    count++;
                }

                // 処理中止かどうかをチェック
                if (waitDlg.IsAborting == true)
                {
                    break;
                }
                else
                {
                    // 進行状況ダイアログのメーターを設定
                    waitDlg.Msg = files.Count.ToString() + "件中: " + count.ToString() + "件目";
                    waitDlg.PerformStep();
                    Application.DoEvents();
                }
            }
            Debug.WriteLine("count = \t" + count);

            // 最終メッセージを表示
            if (waitDlg.DialogResult == DialogResult.Abort)
            {
                waitDlg.Msg = "処理を中断しました。";
            }
            else
            {
                waitDlg.Msg = "処理を完了しました。";
            }
            Application.DoEvents();
            System.Threading.Thread.Sleep(100);
            waitDlg.Close();

            // プログレスバーを表示終了
            splitButton_insert.Enabled = true;
        }


        private void pasteImage(Excel.Worksheet sheet, Excel.Range cell, string imagePath)
        {
            if (checkBox_cell.Checked)
            {
                // セルに画像貼付
                pasteImageOnCell(sheet, cell, imagePath);
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

        private void pasteImageOnCell(Excel.Worksheet sheet, Excel.Range cell, string imagePath)
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
            if (editBox_setW.Enabled)
            {
                cell.ColumnWidth = int.Parse(editBox_setW.Text) * ratioC;
            }
            if (editBox_setH.Enabled)
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
