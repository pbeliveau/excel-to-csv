' Philippe Beliveau        pbeliveau
' MIT Licence
'
' Generate a text file from an Excel sheet.
' Each row becomes a line and each cells are
' separated by the string determined by the
' "strSeparator" value.
'
' VERSION 1.0
'
' Release history :
'      Version 1.0,    June 05, 2018  ~   Initial release

Option Explicit
Public Const strSeparator   As String = ","
Dim strPath                 As String
Dim shtName                 As String
Dim shtData                 As Worksheet
Dim intCol                  As Long
Dim intRow                  As Long

Sub Main()

    Dim fso             As Object
    Dim oFile           As Object
    Dim strFile         As String
    Dim strTime         As String
    Dim strLog          As String
    Dim i               As Long: i = 1
    Dim a               As Long
    
    On Error GoTo ErrHandler
    
    ' Register path and data sheet
    Call regWorksheet
    Call regPath
    intCol = Application.CountA(shtData.Range("1:1"))
    intRow = Application.CountA(shtData.Range("A:A"))
    
    ' Create strFile to log each line
    strFile = ActiveSheet.Name & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(strPath & "\" & strFile)

    ' Create each line and write it to the csv file
    With shtData
        Do While .Cells(i, 1).Value <> ""
            For a = 1 To intCol
                strLog = strLog & _
                         shtData.Cells(i, a).Value & _
                         strSeparator
            Next a
            strLog = Left(strLog, Len(strLog) - 1)
            Call updateBar(Format(i / intRow, "0.0%"))
            i = i + 1
            oFile.WriteLine strLog
            strLog = ""
        Loop
    End With

    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
    Call updateBar("")
    MsgBox "Done."
    Exit Sub

ErrHandler:
    If Err.Number = 9 Then
        MsgBox "The name of the sheet is wrong. Please " & _
                "try again with an adequate sheet name."
        Exit Sub
    Else
        Exit Sub
    End If
        
End Sub
Private Sub regPath()

    Dim selectedFolder
    
    MsgBox "You will now be prompted to select " & _
            "the directory where the text file " & _
            "will be saved."
            
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        strPath = .SelectedItems(1)
    End With
    
End Sub
Private Sub regWorksheet()

    shtName = InputBox("Indicate the name of the " & _
                       "sheet with the data:", _
                       "Name of the sheet", ActiveSheet.Name)
                       
    Set shtData = ActiveWorkbook.Worksheets(shtName)
    
End Sub
Private Sub updateBar(strStatus As String)
    Application.StatusBar = strStatus
End Sub
