Attribute VB_Name = "Module4"
'TO SAVE SINGLE WORKSHEET AS PDF

Sub SinglePDFActiveSheet()

Dim ws As Worksheet
Dim wb As Workbook
Dim strName As String
Dim strPath As String
Dim strFile As String
Dim strPathFile As String
Dim myFile As Variant
Dim strTime As String
On Error GoTo errHandler

Set wb = ActiveWorkbook
Set ws = ActiveSheet
strTime = Format(Now(), "mm yyyy")

'get active workbook folder, if saved
strPath = wb.Path
If strPath = "" Then
  strPath = Application.DefaultFilePath
End If

strPath = strPath & "\"

'replace spaces and periods in sheet name
strName = Replace(ws.Name, " ", "")
strName = Replace(strName, ".", "_")

'create default name for saving file
strFile = strName & "_" & strTime & ".pdf"
strPathFile = strPath & strFile

'export to PDF in current folder
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=strPathFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

exitHandler:
    Exit Sub

errHandler:
    MsgBox "Could not create PDF file"
    
    Resume exitHandler
    
End Sub





