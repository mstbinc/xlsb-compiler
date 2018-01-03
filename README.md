## XLSB Compiler

A program that takes Visual Basic source code and compiles it into a macro-enabled workbook.

When working with VBA modules in macro enabled workbooks, Excel saves the VBA code in a compressed binary format within the rest of the workbook file.
This means that if you want to modify your VBA code, you have to use the VBA IDE that is tied in with Excel.
This is impractical if you plan on using source control to keep track of changes to your code, or if you want to use your own text editor to edit the raw VBA source.

The purpose of this program is to keep the `.vba` source modules separate from the workbook file to grant source control and freedom from the Excel VBA IDE.

### Build Dependencies

This project depends on [Office Primary Interop Assemblies](https://msdn.microsoft.com/en-us/library/15s06t57.aspx) for Excel.

### Running

1. Place all `.vba` source files into `./src/` relative to the program executable.
2. Run the `BinaryConvert.exe`.
3. a `PERSONAL.xlsb` file will be generated.

The newly created file can be copied to `%appdata%\Microsoft\Excel\XLSTART` to run on Excel startup.
VBA modules are named after their file name.

### Assigning Shortcut Keys to Macros

Shortcut keys can be defined for each macro by including the following line your sub procedure block:

`Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "y\n14"`

where `y` is the shortcut key you want to assign. `\n14` specifies the <kbd>Ctrl</kbd> key.
Adding an uppercase shortcut key will change the combination to <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+`key`