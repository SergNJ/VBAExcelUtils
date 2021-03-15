Attribute VB_Name = "CutSCFile"
Option Explicit

Sub CutSAllScorecardsFile()
Dim Path, FileName, CurrCustName As String
Dim wb, wbALLSC, wbSC As Workbook
Dim i, j As Integer

Set wb = ActiveWorkbook
Path = wb.Path 'Current folder path

FileName = Path & "\filename.xlsx"
Set wbALLSC = Workbooks.Open(FileName, 3, True)

i = 1
While i <= wbALLSC.Sheets.Count
        Set wbSC = Workbooks.Add
        
        wbALLSC.Sheets(i).Copy After:=wbSC.Sheets(wbSC.Sheets.Count)
        
        'Deleting unnecessary default sheets in the source file
        Application.DisplayAlerts = False
        j = 1
            While j <= wbSC.Sheets.Count
                If wbSC.Sheets(j).Name = "Sheet1" Or wbSC.Sheets(j).Name = "Sheet2" Or wbSC.Sheets(j).Name = "Sheet3" Then
                wbSC.Sheets(j).Delete
                j = j - 1
                End If
             j = j + 1
            Wend
        Application.DisplayAlerts = True
            
        wbSC.SaveAs (Path & "\Output\output_" & wbSC.Sheets(1).Name & "_" & Format(Now(), "yyyy-MM-dd") & ".xlsx")
        wbSC.Close SaveChanges:=False
        
        Set wbSC = Nothing
        i = i + 1
Wend

wbALLSC.Close SaveChanges:=False
MsgBox i - 1 & " files were created!"
End Sub
