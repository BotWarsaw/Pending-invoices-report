Function main(OpenInvoicePath As String, ParkedInvoicePath As String)
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = wb.Sheets(1)

Dim OpenInv As Workbook: Set OpenInv = Application.Workbooks.Open(OpenInvoicePath)
Dim ParkedInv As Workbook: Set ParkedInv = Application.Workbooks.Open(ParkedInvoicePath)


Dim OpenSh As Worksheet: Set OpenSh = OpenInv.Worksheets(1)
Dim ParkedSh As Worksheet: Set ParkedSh = ParkedInv.Worksheets(1)

'Dim DiffDate  As Date
Dim PostingDate As Date

Dim Report As Worksheet: Set Report = wb.Worksheets.Add

Dim x As Worksheet
For Each x In wb.Worksheets

If x.Name = "Report" Then
x.Delete

End If
Next x
Report.Name = "Report"




Report.Cells(1, 1).Value = "Company Code"
Report.Cells(1, 2).Value = "Document No"
Report.Cells(1, 3).Value = "Amount"
Report.Cells(1, 4).Value = "Posting Date"
Report.Cells(1, 5).Value = "Comment"
Report.Cells(1, 6).Value = "Approver"
Report.Cells(1, 7).Value = "Days No."

pstDate = 0
ReferenceClmn = 0
Amountindoccurr = 0
companyCodeClm= 0

for j = 1 to 50
If ParkedSh.Cells(1, j).Value = "Posting Date" Then
pstDate = j
End If
If ParkedSh.Cells(1, j).Value = "Document Number" Then
ReferenceClmn = j
End If

If ParkedSh.Cells(1, j).Value = "Amount in doc. curr." Then
Amountindoccurr = j
End If

If ParkedSh.Cells(1, j).Value = "Company Code" Then
companyCodeClm= j
End If

next j
 i = 2
Do Until IsEmpty(ParkedSh.Cells(i, companyCodeClm))
CompanyCode = ParkedSh.Cells(i, companyCodeClm ).Value

If Not CompanyCode = "" Then


DocNo = ParkedSh.Cells(i, ReferenceClmn ).Value
PostingDate = ParkedSh.Cells(i, pstDate ).Value
Amount = ParkedSh.Cells(i, Amountindoccurr ).Value

Report.Cells(i, 1).Value = CompanyCode
Report.Cells(i, 2).Value = DocNo
Report.Cells(i, 3).Value = Amount
Report.Cells(i, 4).Value = PostingDate
For j = 1 To 20

If OpenSh.Cells(1, j).Value = "Text" Then
commentColumnOpenSh = j
End If


If OpenSh.Cells(1, j).Value = "Document Number" Then
docNoColumnOpenSh = j
End If 


Next j


match_open = Application.Match(DocNo, OpenSh.Columns(docNoColumnOpenSh ), 0)

        If IsError(match_open) Then
        Report.Cells(i, 5).Value = "Missing in OCR"
        Else
        Report.Cells(i, 5).Value = OpenSh.Cells(match_open, commentColumnOpenSh).Value
        
        End If

CheckEmail = ExtractEmailFun(Report.Cells(i, 5).Value)


If CheckEmail <> "" Then

Report.Cells(i, 6).Value = CheckEmail

        ElseIf Report.Cells(i, 5) = "Missing in OCR" Then
        
        Report.Cells(i, 6).Value = ""
        
        Else

        Report.Cells(i, 6).Value = ""
End If




DiffDate = DateDiff("d", PostingDate, Date)
Report.Cells(i, 7).Value = DiffDate

End If




i = i + 1
Loop




OpenInv.Close
ParkedInv.Close
Report.Columns.AutoFit
Report.Activate

Application.DisplayAlerts = True

Application.ScreenUpdating = True

End Function

Function ExtractEmailFun(extractStr As String) As String
'Update by extendoffice
Dim CharList As String
On Error Resume Next
CheckStr = "[A-Za-z0-9._-]"
OutStr = ""
Index = 1
Do While True
    Index1 = VBA.InStr(Index, extractStr, "@")
    getStr = ""
    If Index1 > 0 Then
        For p = Index1 - 1 To 1 Step -1
            If Mid(extractStr, p, 1) Like CheckStr Then
                getStr = Mid(extractStr, p, 1) & getStr
            Else
                Exit For
            End If
        Next
        getStr = getStr & "@"
        For p = Index1 + 1 To Len(extractStr)
            If Mid(extractStr, p, 1) Like CheckStr Then
                getStr = getStr & Mid(extractStr, p, 1)
            Else
                Exit For
            End If
        Next
        Index = Index1 + 1
        If OutStr = "" Then
            OutStr = getStr
        Else
            OutStr = OutStr & Chr(10) & getStr
        End If
    Else
        Exit Do
    End If
Loop
ExtractEmailFun = OutStr
End Function




Function Pivot()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long

On Error Resume Next
Application.DisplayAlerts = False
ThisWorkbook.Worksheets("Pivot").Delete
Set PSheet = ThisWorkbook.Sheets.Add

PSheet.Name = "Pivot"
Application.DisplayAlerts = True
'Set PSheet = Worksheets("Pivot")
Set DSheet = Worksheets("Report")



LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)



'' FIRST PIVOT
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(5, 1), _
TableName:="CompanyCode_Approvers_Count")


Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="CompanyCode_Approvers_Count")

With PSheet.PivotTables("CompanyCode_Approvers_Count").PivotFields("Company Code")
.Orientation = xlRowField
.Position = 1
End With

With PSheet.PivotTables("CompanyCode_Approvers_Count").PivotFields("Approver")
.Orientation = xlRowField
.Position = 2
End With


With PSheet.PivotTables("CompanyCode_Approvers_Count").PivotFields("Document No")
.Orientation = xlDataField
.Position = 1
.Function = xlCount
.Name = "Number of Documents"
End With


 
    With PSheet.PivotTables("CompanyCode_Approvers_Count").PivotFields( _
        "Approver")
        .PivotItems("NO OCR").Visible = False
        .PivotItems("SSC").Visible = False
    End With
    PSheet.PivotTables("CompanyCode_Approvers_Count").PivotFields("Approver"). _
        AutoSort xlDescending, "Number of Documents"
        
        
        
        

'' SECOND PIVOT
    
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(5, 5), _
TableName:="List of documents")

Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="List of documents")

With PSheet.PivotTables("List of documents").PivotFields("Approver")
.Orientation = xlRowField
.Position = 1
End With


With PSheet.PivotTables("List of documents").PivotFields("Document No")
.Orientation = xlRowField
.Position = 2
End With
With PSheet.PivotTables("List of documents").PivotFields("Days No.")

.Orientation = xlDataField
.Function = xlSum

End With

 
    With PSheet.PivotTables("List of documents").PivotFields( _
        "Approver")
        .PivotItems("NO OCR").Visible = False
        .PivotItems("SSC").Visible = False
    End With
    PSheet.PivotTables("List of documents").PivotFields("Approver"). _
        AutoSort xlDescending, "Number of Documents"







'' THIRD PIVOT
    
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(5, 9), _
TableName:="Approver summary")

Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="Approver summary")

With PSheet.PivotTables("Approver summary").PivotFields("Approver")
.Orientation = xlRowField
.Position = 1
End With


With PSheet.PivotTables("Approver summary").PivotFields("Company Code")
.Orientation = xlRowField
.Position = 2
End With
With PSheet.PivotTables("Approver summary").PivotFields("Days No.")

.Orientation = xlDataField
.Function = xlCount

End With

 
    With PSheet.PivotTables("Approver summary").PivotFields( _
        "Approver")
        .PivotItems("NO OCR").Visible = False
        .PivotItems("SSC").Visible = False
    End With
    PSheet.PivotTables("Approver summary").PivotFields("Approver"). _
        AutoSort xlDescending, "Count of Days No."






'' FOURTH PIVOT
    
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(5, 13), _
TableName:="Vendor Summary")

Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="Vendor Summary")

With PSheet.PivotTables("Vendor Summary").PivotFields("Vendor Name")
.Orientation = xlRowField
.Position = 1
End With


With PSheet.PivotTables("Vendor Summary").PivotFields("Document No")
.Orientation = xlRowField
.Position = 2
End With
With PSheet.PivotTables("Vendor Summary").PivotFields("Days No.")

.Orientation = xlDataField
.Function = xlSum

End With

 
    With PSheet.PivotTables("Vendor Summary").PivotFields( _
        "Vendor Name")
        .PivotItems("(blank)").Visible = False
        
    End With
    PSheet.PivotTables("Vendor Summary").PivotFields("Vendor Name"). _
        AutoSort xlDescending, "Sum of Days No."





    PSheet.Cells(4, 1).Value = "Pivot #1"
    
    PSheet.PivotTables("CompanyCode_Approvers_Count").CompactLayoutRowHeader _
        = "Count Invoices pending approval per CoCode per Approver"
    PSheet.Range("A5").Select
    Selection.Copy
    PSheet.Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    
    
    
    PSheet.Cells(4, 5).Value = "Pivot #2"
    
    PSheet.PivotTables("List of documents").CompactLayoutRowHeader _
        = "List of invoices pending approval per days per Approver"
    PSheet.Range("E5").Select
    Selection.Copy
    PSheet.Range("E4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
        PSheet.Cells(4, 9).Value = "Pivot #3"
    
    PSheet.PivotTables("Approver summary").CompactLayoutRowHeader _
        = "Count of invoices pending approval per Approver"
    PSheet.Range("I5").Select
    Selection.Copy
    PSheet.Range("I4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    
            PSheet.Cells(4, 13).Value = "Pivot #4"
    
    PSheet.PivotTables("Vendor Summary").CompactLayoutRowHeader _
        = "List of invoices pending approval per days overdue. "
    PSheet.Range("M5").Select
    Selection.Copy
    PSheet.Range("M4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

'FORMATTING SECOND PIVOT

lastRowPSheet = PSheet.Cells(Rows.Count, 6).End(xlUp).Row

Dim rng As Range
Set rng = PSheet.Cells(PSheet.Cells(7, 6), PSheet.Cells(lastRowPSheet, 6))



    PSheet.Columns("F:F").Select
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights2)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValuePercent
        .Value = 33
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValuePercent
        .Value = 67
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValuePercent
        .Value = 33
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValuePercent
        .Value = 67
        .Operator = 7
    End With
    Cells.FormatConditions.Delete
    Range("BY7").Select
    Range("F1").Activate
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    Selection.FormatConditions(1).IconCriteria(1).Icon = xlIconGreenCircle
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 3
        .Operator = 7
        .Icon = xlIconYellowCircle
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 5
        .Operator = 7
        .Icon = xlIconRedCircleWithBorder
    End With
    Columns("F:F").Select
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    Selection.FormatConditions(1).IconCriteria(1).Icon = xlIconGreenCircle
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 3
        .Operator = 7
        .Icon = xlIconYellowCircle
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 5
        .Operator = 7
        .Icon = xlIconRedCircleWithBorder
    End With

Application.DisplayAlerts = True

Application.ScreenUpdating = True

End Function