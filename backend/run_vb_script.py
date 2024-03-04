import os
import subprocess

def run_vb_script():
    # VBScript code as a string
    vbscript_code = """Option Explicit

       Dim excelApp, workbook, worksheet, currentDirectory, rowIndex, colIndex

       ' Get the current directory
       currentDirectory = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

       ' Create an Excel application object
       Set excelApp = CreateObject("Excel.Application")

       ' Make Excel visible (optional, you can set it to False if you want it to run in the background)
       excelApp.Visible = False

       ' Open the Excel file in the current directory
       Set workbook = excelApp.Workbooks.Open(currentDirectory & "\Dashboard_.xlsx")

       ' Set a reference to the active sheet
       Set worksheet = workbook.ActiveSheet

       ' Iterate through rows 1 through 1000
       For rowIndex = 1 To 1000
           ' Check if the cell value in column A is 15
           If worksheet.Cells(rowIndex, 1).Value = 15 Then
               ' Delete the entire row
               worksheet.Rows(rowIndex).Delete
               ' Adjust rowIndex to account for the deleted row
               rowIndex = rowIndex - 1
           ' Delete Cell Marker Value
           Else
           worksheet.Cells(rowIndex, 1).Value = ""     
           End If
       Next

       ' Iterate through columns A (1) through BZ (78)
       For colIndex = 1 To 78
           ' Check if the cell value in row 1 is 3
           If worksheet.Cells(1, colIndex).Value = 3 Then
               ' Delete the entire column
               worksheet.Columns(colIndex).Delete
               ' Adjust colIndex to account for the deleted column
               colIndex = colIndex - 1
           ' Delete Cell Marker Value
           Else
           worksheet.Cells(1, colIndex).Value = ""        
           End If
       Next

       ' Save the changes to the Excel file
       workbook.Save

       ' Close Excel
       workbook.Close
       excelApp.Quit

       ' Release the objects
       Set worksheet = Nothing
       Set workbook = Nothing
       Set excelApp = Nothing
       """""


    # Save VBScript code to a temporary file
    vbscript_file_path = "temp.vbs"
    with open(vbscript_file_path, "w") as vbscript_file:
        vbscript_file.write(vbscript_code)

    # Run the VBScript using cscript.exe
    try:
        subprocess.run(["cscript.exe", "//NoLogo", vbscript_file_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running VBScript: {e}")
    finally:
        # Remove the temporary VBScript file
        try:
            os.remove(vbscript_file_path)
        except OSError as e:
            print(f"Error deleting temporary file: {e}")