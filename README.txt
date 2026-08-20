
SHIPMENT TEMPLATE AUTOMATION - PORTABLE WINDOWS BUILD

FINAL TARGET:
The work PC does NOT need Python installed.

The final ZIP contains:
- portable Python runtime
- openpyxl
- pywin32
- Tkinter
- application files
- START_SHIPMENT.bat

The work PC DOES need:
- Microsoft Excel installed
- Windows

No LibreOffice is required.
No Python installation is required.
No pip installation is required on the work PC.
No EXE is required.

WORKFLOW:
1. Build this repository with GitHub Actions.
2. Download the artifact Shipment_Automation_Portable.zip.
3. Extract it on the work PC.
4. Double-click START_SHIPMENT.bat.
5. Select multiple source Excel files.
6. Select the template.
7. Choose output folder.
8. Run.

IMPORTANT TEMPLATE RULE:
Row 2 of ShipmentDeatils and ShipmentItemsDetails is the master template row.
Every value and every formula present in row 2 is copied/fill-down to each
generated row. Formulas are translated to the destination row.
Formatting, borders, number formats, alignment and protection are also copied.
Only explicitly source/business-controlled fields are overwritten.

EXCEL FORMAT:
Sources may be .xls/.xlsx/.xlsm.
Final generated shipment file is legacy .xls.
Microsoft Excel COM is used for .xls conversion and final .xls Save As.

VALIDATION:
Unknown commodity/HS-code mappings stop the job instead of guessing.
Source files and the original template are never modified.

PREVIOUS FBT LESSONS BUILT IN:
- Excel COM, not LibreOffice
- test the win32com package during build and test actual Excel COM when the application runs
- complete row-2 donor captured BEFORE clearing rows
- GUI arguments match processor function
- no undefined helper calls
- temporary conversion files are isolated and deleted
- final output uses Excel SaveAs .xls


GITHUB ACTIONS NOTE:
The GitHub-hosted Windows builder does not have Microsoft Excel installed.
Therefore the workflow verifies Python/Tkinter/openpyxl/pywin32 and the
application import, but does not falsely claim that Excel COM was tested
on GitHub. The application performs a real Excel COM availability check
on the user's work PC immediately before processing.


TKINTER/TCL NOTE:
The portable runtime must keep its Tcl/Tk files. Do not remove runtime\tcl,
because Tkinter requires init.tcl at runtime.
