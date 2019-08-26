# Visual Basic Upgrade Companion Version 7.2

# Release Notes

# Visual Studio 2017
VBUC version 7.2 generates Visual Studio 2017 compliant source code, as well as project and solution files.

You can now choose what kind of VS solution file you want to generate up to and including VS 2017 (version 15).

# Typing Inference

## Robustness and Scalability:

The typing inference mechanism is faster and uses memory more efficiently.

Some projects that
could not be correctly analyzed with previous versions of VBUC can now be fully analyzed and migrated, achieving significant improvements in the quality of the migrated code.

## Multidimensional Arrays
In VB6, arrays can mutate and change dimensions and magnitudes, while in .Net they must be declared in the simplest way that will be able to represent all their previous behaviors.

The current typer version identifies several new patterns related to the multiple dimensional arrays and makes much better decisions regarding this area.

## COM Let/Set Properties

`Let` and `Set` usage of COM properties has been improved to produce a correct typing analysis in those cases for parameters and return types.

# `R-Value` vs `L-Value` Influences
The heuristics for how expressions in an assignment influence each other has been modified to recognize more influence from the right to the left side. Expressions appearing at the left side of
an assignment must be able to take the type of the right-side expression. That doesn´t work in the same way in the opposite direction.

# Other Bugs
Several other bugs were fixed in the VBUC typing inference.

# Helpers Enhancements
Multiple helper classes have been updated to improve performance and stability, while other new
classes have been added to support more VB6 functionality, among them:

- FPSreapd helper: Farpoint Spread Control support class
- ADORecordsetHelper: support class for the ADO.Recordset
- PowerPacks helper: supports Visual Basic Printing Library
- PictureBoxExtended helper: supports advanced PictureBox functionality in VB6.
- 
# Increased Mappings Coverage
A significant amount of preexisting third-party control mappings have been updated to fix
compilation issues, improve performance, or increase functionality coverage. Additionally, many
new mappings have been implemented in order to support more third party libraries that include
commonly used controls, for example:
- Accusoft option, to support common classes for this library
- CRAXDRT_CRVIEWERLibCtl option, which includes mapping for Crystal Report Libraries
- CWUIControlsLib option, to support controls of National Instruments Library
- MemoLibfpMemo option, which includes some common control of Farpoint libraries to
standard .Net controls
- OracleInProc option, which helps map OracleInProc server database libraries to .Net Data
  
# Connection classes.
Parallel Migration - Improved Stability
VBUC 7.1 included the first version of the VBUC Parallel Migration option, which, using spare cores,
executes the upgrade step of each project in the migration solution concurrently, reducing overall
processing time dramatically.
VBUC version 7.2 includes a more stable version of this feature, tested over millions of lines of realworld code and several specialized stress cases.

# Other improvements
There are between 100 to 200 individual fixes related to areas such as:
- class implementations,
- COM Interop,
- type coercions,
- operations,
- parameter passing,
- reference qualifications,
- structures conversion,
- arrays, 
- and many more.

This newest version continues to reduce the errors that block compiling the migrated code, thus saving time in the final stages of the project.
A few other issues where fixed for the latest version of the assessment tool, both related to counts and to reports format issues.
 