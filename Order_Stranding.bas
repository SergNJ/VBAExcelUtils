Attribute VB_Name = "Order_Stranding"
' SERGIY PRYKHODKO
' THERMO FISHER SCIENTIFIC
' sergiy.prykhodko@thermofisher.com
' 201-469-5677

Sub OS_Show_FPAK_Only()
Attribute OS_Show_FPAK_Only.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As String
Dim wb As Workbook

'
' OS_Show_FPAK_Only Macro
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("ORDERS").Activate
    s = ActiveCell.Address
    ActiveSheet.Range("$G$3:$FY$10000").AutoFilter Field:=5, Criteria1:="YES"
    Range(s).Select
    Application.ScreenUpdating = True

End Sub
Sub OS_Show_All_Orders()
Attribute OS_Show_All_Orders.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As String
Dim wb As Workbook

'
' OS_Show_All_Orders Macro
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("ORDERS").Activate
    s = ActiveCell.Address
    Range("$G$3:$FY$10000").Select
    Selection.AutoFilter
    Range(s).Select
    Application.ScreenUpdating = True
    
End Sub

Sub OS_FPak_Orders_with_No_ID()
Dim s As String
Dim wb As Workbook
'
' OS_FPak_Orders_with_No_ID Macro
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("ORDERS").Activate
    s = ActiveCell.Address
    ActiveSheet.Range("$G$3:$FY$10000").AutoFilter Field:=6, Criteria1:="YES"
    Range(s).Select
    Application.ScreenUpdating = True
    
End Sub

Sub OS_Add_Grace_Period()
Dim Order, Cust, PN, DID As String
Dim i, r As Integer
Dim wb As Workbook
Dim wsOrd, wsFlags As Worksheet

'
' Add_Grace_Period Macro
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    
    If wb.ActiveSheet.Name <> "ORDERS" Then
        MsgBox "You have to be on the tab ORDERS in order to use this function."
        Exit Sub
    Else:
        Set wsOrd = wb.Worksheets("ORDERS")
        Set wsFlags = wb.Worksheets("FLAGS")
    End If
    
    
    r = ActiveCell.Row
    Order = wsOrd.Range("S" & r).Value
    Cust = wsOrd.Range("O" & r).Value
    PN = wsOrd.Range("I" & r).Value
    DID = wsOrd.Range("M" & r).Value
    
    i = 2
    For i = 2 To 10000

        If wsFlags.Range("A" & i).Value = "" Then
            wsFlags.Range("A" & i).Value = Order
            wsFlags.Range("B" & i).Value = 1
            Exit For
        End If

Next i

     
    Application.ScreenUpdating = True
    MsgBox "Grace period was added to the drum #" & DID & ", " & PN & " under " & Order & " from " & Cust
    
End Sub

Sub OS_Show_All_Drums()
Dim s As String
Dim wb As Workbook

'
' OS_Show_All_Drums
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("RDS").Activate
    s = ActiveCell.Address
    Range("$A$3:$S$7002").Select
    Selection.AutoFilter
    Range(s).Select
    Application.ScreenUpdating = True
    
End Sub
Sub OS_NJ_Empties_No_Ord()
Dim s As String
Dim wb As Workbook

'
' OS_NJ_Empties_No_Ord
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("RDS").Activate
    s = ActiveCell.Address
    ActiveSheet.Range("$A$3:$S$7002").AutoFilter Field:=6, Criteria1:="YES"
    Range(s).Select
    Application.ScreenUpdating = True
    
End Sub
Sub OS_Ord_Nonproductive()
Dim s As String
Dim wb As Workbook

'
' OS_NJ_Empties_No_Ord
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("RDS").Activate
    s = ActiveCell.Address
    ActiveSheet.Range("$A$3:$S$7002").AutoFilter Field:=7, Criteria1:="YES"
    Range(s).Select
    Application.ScreenUpdating = True
    
End Sub
Sub OS_Duplicated_Ord()
Dim s As String
Dim wb As Workbook

'
' OS_Duplicated_Ord Macro
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("ORDERS").Activate
    s = ActiveCell.Address
    
    ActiveSheet.Range("$G$3:$FY$10000").AutoFilter Field:=23, Criteria1:="YES"
    ActiveWorkbook.Worksheets("ORDERS").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("ORDERS").AutoFilter.Sort.SortFields.Add Key:=Range("M3:M10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ORDERS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(s).Select
    Application.ScreenUpdating = True

End Sub
Sub OS_Fill_Scheduling()
Dim s As String
Dim wb As Workbook

'
' OS_Fill_Scheduling
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    wb.Worksheets("ORDERS").Activate
    s = ActiveCell.Address
    
    ActiveSheet.Range("$G$3:$FY$10000").AutoFilter Field:=5, Criteria1:="YES"
    ActiveSheet.Range("$G$3:$FY$10000").AutoFilter Field:=6, Criteria1:="NO"
    ActiveSheet.Range("$G$3:$FY$10000").AutoFilter Field:=8, Criteria1:="YES"
    
    ActiveWorkbook.Worksheets("ORDERS").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ORDERS").AutoFilter.Sort.SortFields.Add Key:=Range("H4:H10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ORDERS").AutoFilter.Sort.SortFields.Add Key:=Range("V4:V10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ORDERS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(s).Select
    Application.ScreenUpdating = True
    
End Sub


Sub OS_Flag_As_Processed()
Dim Order, Cust, PN, DID As String
Dim i, r As Integer
Dim wb As Workbook
Dim wsOrd, wsFlags As Worksheet

'
' OS_Flag_As_Processed Macro
'

'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    
    If wb.ActiveSheet.Name <> "ORDERS" Then
        MsgBox "You have to be on the tab ORDERS in order to use this function."
        Exit Sub
    Else:
        Set wsOrd = wb.Worksheets("ORDERS")
        Set wsFlags = wb.Worksheets("FLAGS")
    End If
    
    
    r = ActiveCell.Row
    Order = wsOrd.Range("S" & r).Value
    Cust = wsOrd.Range("O" & r).Value
    PN = wsOrd.Range("I" & r).Value
    DID = wsOrd.Range("M" & r).Value
    
    i = 2
    For i = 2 To 10000

        If wsFlags.Range("A" & i).Value = "" Then
            wsFlags.Range("A" & i).Value = Order
            wsFlags.Range("C" & i).Value = 1
            Exit For
        End If

Next i

     
    Application.ScreenUpdating = True
    MsgBox "This order was flagged as PROCESSED: drum #" & DID & ", " & PN & " under " & Order & " from " & Cust
    
End Sub
