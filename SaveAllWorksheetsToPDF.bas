Attribute VB_Name = "Module3"
'TO SAVE ALL THE WORKSHEETS TO PDF


Sub WorkSheetsToPDF()

Dim ws As Worksheet
Dim i As Long
Dim nm As String

'To skips first 2 sheets, put i=3 (For this workbook, skip the first two worksheets "INSTRUCTIONS" and "DATASHEET", hence i=2)
'Hide the tabs you do not want to convert to PDF

For i = 3 To ThisWorkbook.Worksheets.Count

    If Sheets(i).Visible <> xlSheetVisible Then  'This line of code will not convert HIDDEN tabs to pdf
    
    Else
        With ThisWorkbook.Worksheets(i)
        nm = .Name
         'Change the FILEPATH and FOLDER NAME in the filename below:
        .ExportAsFixedFormat Type:=xlTypePDF, _
                                    Filename:="R:\Backup Servicing Clients\_Backup Servicing\2018 Verification Invoicing\testing\06_2018 Verification Invoice_" & nm & ".pdf", _
                                    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                                    IgnorePrintAreas:=True, OpenAfterPublish:=False
        End With
     End If
     
Next i
   
MsgBox "All files converted to PDF successfully"
   
End Sub



