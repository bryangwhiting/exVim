# exVim
An Excel add-in that enables vi/vim-like keystrokes.

# What it does
Turn on this add-in and you can navigate Excel using <kbd>h</kbd>, <kbd>j</kbd>, <kbd>k</kbd>, <kbd>l</kbd> keys. Additionally, within each each keystroke is embedded a different macro. For example, <kbd>y</kbd> will copy text, while <kbd>Y</kbd> will duplicate the sheet. <kbd>d</kbd> deletes the row and <kbd>D</kbd> deletes the column. Use <kbd>i</kbd> or <kbd>a</kbd> to edit the cell (equivalent to <kbd>F2</kbd>). Only numbers 1-9 and the <kbd>=</kbd> keys are non-binded, meaning if you type 1-9, you'll just enter data into the active cell.

Most of the time, running a VBA macro disables your ability to "undo". This add-in has kept such a problem in mind and tries to mostly use macros that don't override your ability to "undo". (Meaning the VBA "Application.SendKeys" function is used widely.)

# How to Install
I could have just put the Excel add-in here for download, but I thought it would be better if users self-installed the add-in, so that they can read what's inside. Follow these steps to install this add-in.
  1. Create a new workbook. Save it as a macro-enabled workbook, i.e. "exVim.xlsm".
  2. Open VBA. [Developer > Visual Basic].
    3. If the Developer ribbon is not active, enable it. [File > Options > Customize Ribbon. Then click the checkbox next to "Developer"]
  4. Open [VBAProject (exVim.sxlm) > Microsofte Excel > ThisWorkbook]
    5. If Project Explorer is not open, step (4) won't make any sense. Go to [View > Project Explorer (ctrl + R)] to enable.
  6. Copy the text from "ThisWorkbook.txt" into ThisWorkbook. This is the list of keybindings.
  5. Next, you need to create a VBA module that stores the code for the macros used by the keybindings. To do so, within VBA go to [Insert > Module].
  6. Copy the code from "module_exVim_macros" into the script window under the newly-created Module1.

Now all of the macros are embedded in "exVim.xlsm". We'll need to convert "exVim.xlsm" into "exVim.xlam", the add-in form.
  1. First, give the add-in a name. [File > Info > Properties > Title]. Call it "~ exVim". (This is the name it will appear as when you add the Excel add-in.)
  2. Create the add-in. [File > Save As]. (Pick any destination, doesn't matter.) When saving the file, [Save as type > Excel Add-in (*.xlam)]. Save with filename "~exVim.xlam".

The exVim add-in is now saved. To enable the add-in, save and close the workbook "exVim.xlsm". Open any other workbook. Go to [Developer > Excel Add-ins] (the icon with the gears). Click the checkbox next to "~ exVim".

* If "~ exVim" doesn't show up as an add-in, click [Browse]. Click "~exVim.xlam" to load.

# Keystrokes
Coming soon...
