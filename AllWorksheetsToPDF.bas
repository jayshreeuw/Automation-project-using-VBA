Attribute VB_Name = "Module1"
'TO RUN THIS CODE, FIRST UNHIDE THE HIDDEN TABS

Sub AllFilestoPDF()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    
    'Change the PATH NAME and FOLDER NAME in Filename below:
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:="R:\Backup Servicing Clients\_Backup Servicing\2018 Verification Invoicing\testing\06_2018_Verification Invoice_" & ws.Name

Next

End Sub

