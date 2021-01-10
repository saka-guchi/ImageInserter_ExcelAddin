<div align="left">
<img src="./images/icon.png" alt="icon" title="icon" width="10%">
<h1>ImageInserter for Excel Addin</h1>
</div>

- [日本語版](https://qiita.com/saka-guti/items/5fc67c76e42fe95d9f2d)

This is an Excel add-in that inserts images into cells and notes.

## Overview
When using Excel, it takes time to paste the image to fit the cell and resize it.This add-in can optimize the size of cells and images and arrange them as you want in a short time.

<div align="center">
<img src="./images/demo_description.png" alt="description" title="demo_description_JP">
</div>

## Demo

**The explanation image is in Japanese, but in reality it will be in English depending on the environment.**

### Insert images
<div align="center">
<img src="./images/ja/demo_insert.gif" alt="insert" title="insert">
</div>

### Delete images
<div align="center">
<img src="./images/ja/demo_delete.gif" alt="delete" title="delete">
</div>

### Select target (Cell or Memo)
<div align="center">
<img src="./images/ja/demo_general.gif" alt="select target" title="select target">
</div>

### Specify cell size
<div align="center">
<img src="./images/ja/demo_set_cell_size.gif" alt="specify cell size" title="specify cell size">
</div>

### Specifying the storage method in the cell
<div align="center">
<img src="./images/ja/demo_store_in_cell.gif" alt="specify storage method" title="specify storage method">
</div>

### Specifying the information to write to the cell
<div align="center">
<img src="./images/ja/demo_write_to_cell.gif" alt="specifying info" title="specifying info">
</div>

### Specifying the delete method in the cell
<div align="center">
<img src="./images/ja/demo_delete_cell.gif" alt="specifying delete method" title="specifying delete method">
</div>

### Specifying how to align cells
<div align="center">
<img src="./images/ja/demo_arrange_cells.gif" alt="specifying align" title="specifying align">
</div>

### Specify maximum memo size
<div align="center">
<img src="./images/ja/demo_set_memo_size.gif" alt="specify max memo size" title="specify max memo size">
</div>

## Environment

.NET Framework 4.8

### Note

1. After processing with this add-in, "Undo (Ctrl + Z)" operation is not possible. Please save the file in advance.
1. If you select "Fit to cell height" as the storage method, the width of the cell and the image may not match.

## Install

1. Double-click "setup.msi" to install
1. Open Excel and check that "Insert Image" is displayed on the ribbon.

\* If not displayed
1. Open "File> Options> Ribbon Preferences"
1. Check "Insert image"

<div align="center">
<img src="./images/ja/demo_install.gif" alt="install" title="install">
</div>

## Uninstall

1. Open "Control Panel> Apps"
1. Select "ImageInserter_ExcelAddin" and click "Uninstall"

## License

[MIT](./LICENSE)
