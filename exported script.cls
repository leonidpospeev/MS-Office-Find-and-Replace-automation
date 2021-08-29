VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub RunContractFileAssembly()
    
    ' Catching errors
    On Error GoTo catch_exception

    ' Getting the file with Excel data
    Dim ExcelFileWithInitData As String
    ExcelFileWithInitData = ActiveWorkbook.Path + "\" + Cells(1, 2).Value + ".xlsx"

    ' Getting the name of the Word file to save results to
    Dim WordFileToStartFrom As String
    WordFileToStartFrom = ActiveWorkbook.Path + "\" + Cells(2, 2).Value + ".docx"
    
    ' A dictionary of keys and values from the Excel file
    Dim Dict
    Set Dict = CreateObject("Scripting.Dictionary")
    
    ' Iteration variable pointing to the row number in the Excel file
    ' to start building the dictionary with
    Dim i As Integer
    i = 1
    
    ' Opening the Excel file
    Set wb = Workbooks.Open(ExcelFileWithInitData)
    
    ' Filling the dictionary
    With wb.ActiveSheet
        While Not .Cells(i, 1).Value = ""
            Dict.Add .Cells(i, 1).Value, .Cells(i, 2).Value
            i = i + 1
        Wend
    End With
    
    ' Closing the Excel file
    wb.Close
    
    ' Opening the Word file
    Set wordObj = CreateObject("Word.Application")
    Set wordDoc = wordObj.Documents.Open(WordFileToStartFrom)
    
    ' Doing "Find and replace"
    For Each varKey In Dict.Keys()
        wordDoc.Content.Find.Execute FindText:=varKey, _
            ReplaceWith:=Dict(varKey), Replace:=wdReplaceAll
    Next
    
    ' Opening a save file prompt to save the Word file
    Dim SaveDialogOutput As Variant
    
    ' Saving the Word file
    SaveDialogOutput = Application.GetSaveAsFilename(FileFilter:= _
        "Word Files (*.docx), *.docx", Title:="Save MS Word file", _
        InitialFileName:=ActiveWorkbook.Path)
        
    ' If the save prompt not cancelled, save the Word file
    If SaveDialogOutput <> False Then
        wordDoc.SaveAs2 Filename:=SaveDialogOutput
    End If
    
    ' Closing and freeing the memory
    wordDoc.Close
    wordObj.Quit
    
    Set wordObj = Nothing
    Set wordDoc = Nothing
    Set Dict = Nothing
    
    Exit Sub
    
    ' Message box if an exception occured
catch_exception:
    MsgBox ("An error occured during the execution")
        
End Sub
