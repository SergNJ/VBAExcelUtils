Attribute VB_Name = "Open_All_Linked_Files"
Sub Open_All_Linked_Files()
Attribute Open_All_Linked_Files.VB_ProcData.VB_Invoke_Func = " \n14"
Dim LinkList As Variant
Dim i As Integer

LinkList = ActiveWorkbook.LinkSources(xlExcelLinks)
 
If Not IsEmpty(LinkList) Then
       
       For i = LBound(LinkList) To UBound(LinkList)
            Application.Workbooks.Open FileName:=LinkList(i)
       On Error Resume Next
       Next i
 
End If


End Sub
