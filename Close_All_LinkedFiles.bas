Attribute VB_Name = "Close_All_LinkedFiles"
Option Explicit

Sub Close_All_Linked_Files()
Dim LinkList As Variant
Dim i As Integer
Dim wb As Workbook
Dim s As String

Application.DisplayAlerts = True
LinkList = ActiveWorkbook.LinkSources(xlExcelLinks)

If Not IsEmpty(LinkList) Then
       For i = LBound(LinkList) To UBound(LinkList)
            s = Right(LinkList(i), Len(LinkList(i)) - InStrRev(LinkList(i), "\"))
            Set wb = Application.Workbooks(s)
            wb.Activate
            wb.Close SaveChanges:=False
       On Error Resume Next
Next i
 
End If

End Sub
