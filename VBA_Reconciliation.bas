Attribute VB_Name = "Module3"
' ============================================================
' MODULE: Ultimate Reconciliation (Merged Version)
' Engine: Advanced 3-Part Key & Duplicate Handling
' Layout: Side-by-Side Amounts with GL Account & Color Formatting
' ============================================================
Option Explicit

Public Sub RunReconciliationUltimate()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Reconciliation running..."

    Dim wb As Workbook: Set wb = ThisWorkbook

    ' -- 1. Validate required source sheets exist --
    If Not SheetExists(wb, "General_Ledger") Or Not SheetExists(wb, "Vendor_Ledger") Then
        MsgBox "Ensure sheets 'Vendor_Ledger' and 'General_Ledger' exist exactly as named.", vbCritical
        GoTo CleanUp
    End If

    Dim wsGL As Worksheet: Set wsGL = wb.Sheets("General_Ledger")
    Dim wsVL As Worksheet: Set wsVL = wb.Sheets("Vendor_Ledger")

    ' -- 2. Auto-create Formatted Settings sheet --
    Dim wsSettings As Worksheet
    If Not SheetExists(wb, "Settings") Then
        Set wsSettings = wb.Sheets.Add(Before:=wb.Sheets(1))
        wsSettings.Name = "Settings"
        SetupSettingsSheet wsSettings
    Else
        Set wsSettings = wb.Sheets("Settings")
    End If

    Dim tol As Double
    tol = SafeDbl(wsSettings.Range("B3").Value)

    ' -- 3. Prepare Report Sheet --
    Dim wsRes As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next: wb.Sheets("Reconciliation_Report").Delete: On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    Set wsRes = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsRes.Name = "Reconciliation_Report"

    ' Write Headers (12 Columns)
    Dim hdr As Variant
    hdr = Array("Posting Date", "Doc No", "Vendor ID", "Vendor Name", "Ref No", "Tran Type", _
                "GL Account", "Vendor Amount", "GL Amount", "Variance", "Status", "Notes")
    
    Dim col As Long
    For col = 0 To UBound(hdr)
        wsRes.Cells(1, col + 1).Value = hdr(col)
    Next col
    FormatHeader wsRes, UBound(hdr) + 1

    ' -- 4. Dynamic Column Detection --
    Dim glColVendorID As Long, glColRefNo As Long, glColTranType As Long, glColGLAcc As Long
    Dim glColNetAmt As Long, glColPostDate As Long, glColDocNo As Long, glColVendorName As Long
    Dim glLastCol As Long
    
    glLastCol = wsGL.Cells(1, wsGL.Columns.Count).End(xlToLeft).Column
    glColPostDate = FindColumn(wsGL, 1, glLastCol, "PostingDate")
    glColDocNo = FindColumn(wsGL, 1, glLastCol, "DocNo")
    glColVendorID = FindColumn(wsGL, 1, glLastCol, "VendorID")
    glColVendorName = FindColumn(wsGL, 1, glLastCol, "VendorName")
    glColRefNo = FindColumn(wsGL, 1, glLastCol, "RefNo")
    glColTranType = FindColumn(wsGL, 1, glLastCol, "TranType")
    glColNetAmt = FindColumn(wsGL, 1, glLastCol, "NetAmount")
    glColGLAcc = FindColumn(wsGL, 1, glLastCol, "GLAccount") ' <-- Added GL Account Detection

    ' Vendor Ledger Columns
    Dim vlColVendorID As Long, vlColRefNo As Long, vlColTranType As Long, vlColAmount As Long
    Dim vlColPostDate As Long, vlColDocNo As Long, vlColVendorName As Long, vlLastCol As Long
    
    vlLastCol = wsVL.Cells(1, wsVL.Columns.Count).End(xlToLeft).Column
    vlColPostDate = FindColumn(wsVL, 1, vlLastCol, "PostingDate")
    vlColDocNo = FindColumn(wsVL, 1, vlLastCol, "DocNo")
    vlColVendorID = FindColumn(wsVL, 1, vlLastCol, "VendorID")
    vlColVendorName = FindColumn(wsVL, 1, vlLastCol, "VendorName")
    vlColRefNo = FindColumn(wsVL, 1, vlLastCol, "RefNo")
    vlColTranType = FindColumn(wsVL, 1, vlLastCol, "TranType")
    vlColAmount = FindColumn(wsVL, 1, vlLastCol, "Amount")

    ' Validate critical columns
    If glColVendorID = 0 Or vlColVendorID = 0 Or glColNetAmt = 0 Or vlColAmount = 0 Then
        MsgBox "Critical columns missing. Ensure VendorID, RefNo, and Amounts exist.", vbCritical
        GoTo CleanUp
    End If

    ' -- 5. Index Vendor Ledger (Handling Duplicates safely) --
    Dim dictVL As Object: Set dictVL = CreateObject("Scripting.Dictionary")
    Dim vlUsed As Object: Set vlUsed = CreateObject("Scripting.Dictionary")
    dictVL.CompareMode = vbTextCompare
    
    Dim i As Long, sKey As String, vlRows As Long, glRows As Long
    vlRows = wsVL.Cells(wsVL.Rows.Count, vlColVendorID).End(xlUp).Row
    glRows = wsGL.Cells(wsGL.Rows.Count, glColVendorID).End(xlUp).Row

    For i = 2 To vlRows
        sKey = MakeKey(wsVL.Cells(i, vlColVendorID).Value, wsVL.Cells(i, vlColRefNo).Value, wsVL.Cells(i, vlColTranType).Value)
        If Not dictVL.Exists(sKey) Then Set dictVL(sKey) = New Collection
        dictVL(sKey).Add i
    Next i

    ' -- 6. Reconcile Loop --
    Dim outRow As Long: outRow = 2
    Dim glAmt As Double, vlAmt As Double, diff As Double, bestDiff As Double
    Dim bestVLRow As Long, vlRow As Long
    Dim matchCount As Long, mismatchCount As Long, totalVariance As Double
    Dim sStatus As String, sNotes As String, strGLAcc As String

    For i = 2 To glRows
        sKey = MakeKey(wsGL.Cells(i, glColVendorID).Value, wsGL.Cells(i, glColRefNo).Value, wsGL.Cells(i, glColTranType).Value)
        glAmt = SafeDbl(wsGL.Cells(i, glColNetAmt).Value)
        
        If glColGLAcc > 0 Then strGLAcc = SafeStr(wsGL.Cells(i, glColGLAcc).Value) Else strGLAcc = ""
        
        bestVLRow = 0
        bestDiff = 1E+30
        vlAmt = 0
        sNotes = ""

        ' Search for Best Match in Vendor Ledger
        If dictVL.Exists(sKey) Then
            Dim colVL As Collection: Set colVL = dictVL(sKey)
            Dim j As Long
            For j = 1 To colVL.Count
                vlRow = colVL(j)
                If Not vlUsed.Exists(CStr(vlRow)) Then
                    Dim tempVlAmt As Double
                    tempVlAmt = SafeDbl(wsVL.Cells(vlRow, vlColAmount).Value)
                    If Abs(glAmt - tempVlAmt) < bestDiff Then
                        bestDiff = Abs(glAmt - tempVlAmt)
                        bestVLRow = vlRow
                        vlAmt = tempVlAmt
                    End If
                End If
            Next j
        End If

        diff = vlAmt - glAmt ' Variance calculation

        ' Evaluate Match Status
        If bestVLRow > 0 Then
            vlUsed(CStr(bestVLRow)) = True
            If Abs(diff) <= tol Then
                sStatus = "Match"
                matchCount = matchCount + 1
            Else
                sStatus = "Amount Mismatch"
                sNotes = "Difference exceeds tolerance limit."
                mismatchCount = mismatchCount + 1
                totalVariance = totalVariance + diff
            End If
        Else
            sStatus = "Missing in Vendor Ledger"
            diff = 0 - glAmt
            mismatchCount = mismatchCount + 1
            totalVariance = totalVariance + diff
        End If

        ' Write Row
        WriteMergedRow wsRes, outRow, wsGL.Cells(i, glColPostDate).Value, SafeStr(wsGL.Cells(i, glColDocNo).Value), _
            SafeStr(wsGL.Cells(i, glColVendorID).Value), SafeStr(wsGL.Cells(i, glColVendorName).Value), _
            SafeStr(wsGL.Cells(i, glColRefNo).Value), SafeStr(wsGL.Cells(i, glColTranType).Value), _
            strGLAcc, vlAmt, glAmt, diff, sStatus, sNotes
            
        outRow = outRow + 1
    Next i

    ' -- 7. Write Remaining Unmatched Vendor Rows --
    For i = 2 To vlRows
        If Not vlUsed.Exists(CStr(i)) Then
            vlAmt = SafeDbl(wsVL.Cells(i, vlColAmount).Value)
            diff = vlAmt - 0
            
            WriteMergedRow wsRes, outRow, wsVL.Cells(i, vlColPostDate).Value, SafeStr(wsVL.Cells(i, vlColDocNo).Value), _
                SafeStr(wsVL.Cells(i, vlColVendorID).Value), SafeStr(wsVL.Cells(i, vlColVendorName).Value), _
                SafeStr(wsVL.Cells(i, vlColRefNo).Value), SafeStr(wsVL.Cells(i, vlColTranType).Value), _
                "", vlAmt, 0, diff, "Missing in General Ledger", ""
                
            mismatchCount = mismatchCount + 1
            totalVariance = totalVariance + diff
            outRow = outRow + 1
        End If
    Next i

    ' -- 8. Polish and Format (Red/Green Numbers) --
    If outRow > 2 Then
        ' Format Amounts: Green for positive, Red for negative
        wsRes.Range("H2:J" & outRow - 1).NumberFormat = "[Color10]#,##0.00;[Red]-#,##0.00;0.00"
        wsRes.Range("A1:L" & outRow - 1).Borders.LineStyle = xlContinuous
        wsRes.Range("A1:L" & outRow - 1).AutoFilter
    End If
    
    wsRes.Columns.AutoFit
    wsRes.Activate
    wsRes.Range("A1").Select

    ' -- Summary Message Box --
    Dim msg As String
    msg = "Ultimate Reconciliation Complete!" & vbCrLf & vbCrLf & _
          "Tolerance Limit Used: " & tol & vbCrLf & _
          "--------------------------------" & vbCrLf & _
          "Total Matches: " & matchCount & vbCrLf & _
          "Total Mismatches/Missing: " & mismatchCount & vbCrLf & _
          "Total Variance: " & Format(totalVariance, "#,##0.00")
          
    MsgBox msg, vbInformation, "Reconciliation Summary"

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' ----------------------------------------------------
' HELPER FUNCTIONS
' ----------------------------------------------------

Private Function SheetExists(wb As Workbook, shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Sheets(shName): On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function FindColumn(ws As Worksheet, hdrRow As Long, lastCol As Long, hdrName As String) As Long
    Dim c As Long
    FindColumn = 0
    For c = 1 To lastCol
        If InStr(1, UCase(Trim(CStr(ws.Cells(hdrRow, c).Value))), UCase(Trim(hdrName)), vbTextCompare) > 0 Then
            FindColumn = c: Exit Function
        End If
    Next c
End Function

Private Function MakeKey(vID As Variant, ref As Variant, tran As Variant) As String
    MakeKey = UCase(Trim(CStr(vID))) & "|" & UCase(Trim(CStr(ref))) & "|" & UCase(Trim(CStr(tran)))
End Function

Private Function SafeDbl(v As Variant) As Double
    On Error Resume Next: SafeDbl = 0
    If Not IsEmpty(v) And Not IsError(v) Then
        If IsNumeric(v) Then SafeDbl = CDbl(v) Else SafeDbl = Val(CStr(v))
    End If
    On Error GoTo 0
End Function

Private Function SafeStr(v As Variant) As String
    On Error Resume Next: SafeStr = ""
    If Not IsEmpty(v) And Not IsError(v) Then SafeStr = CStr(v)
    On Error GoTo 0
End Function

Private Sub FormatHeader(ws As Worksheet, nCols As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))
        .Interior.Color = RGB(200, 220, 240) ' Light Blue
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub WriteMergedRow(ws As Worksheet, r As Long, pDate As Variant, dNo As String, vID As String, _
    vName As String, ref As String, tType As String, glAcc As String, vAmt As Double, glAmt As Double, _
    diff As Double, status As String, notes As String)
    
    If IsDate(pDate) Then
        ws.Cells(r, 1).Value = CDate(pDate)
        ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd"
    Else
        ws.Cells(r, 1).Value = pDate
    End If
    
    ws.Cells(r, 2).Value = dNo
    ws.Cells(r, 3).Value = vID
    ws.Cells(r, 4).Value = vName
    ws.Cells(r, 5).Value = ref
    ws.Cells(r, 6).Value = tType
    ws.Cells(r, 7).Value = glAcc
    
    ' Leave zeros blank if one side is missing to make it cleaner
    If status = "Missing in Vendor Ledger" Then
        ws.Cells(r, 8).Value = ""
    Else
        ws.Cells(r, 8).Value = vAmt
    End If
    
    If status = "Missing in General Ledger" Then
        ws.Cells(r, 9).Value = ""
    Else
        ws.Cells(r, 9).Value = glAmt
    End If
    
    ws.Cells(r, 10).Value = diff
    ws.Cells(r, 11).Value = status
    ws.Cells(r, 12).Value = notes
End Sub

Private Sub SetupSettingsSheet(ws As Worksheet)
    ws.Cells(1, 1).Value = "Vendor GL Reconciliation - Settings"
    ws.Cells(1, 1).Font.Size = 14: ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Color = RGB(0, 51, 102)
    ws.Cells(3, 1).Value = "Tolerance Limit (Amount):": ws.Cells(3, 1).Font.Bold = True
    
    With ws.Cells(3, 2)
        .Value = 0#
        .Interior.Color = RGB(200, 255, 200) ' Light Green
        .Font.Color = vbBlue: .Font.Bold = True: .NumberFormat = "0.00"
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Cells(3, 3).Value = "<-- Change this value and re-run the macro"
    ws.Cells(3, 3).Font.Italic = True: ws.Cells(3, 3).Font.Color = RGB(128, 128, 128)
    
    ws.Cells(5, 1).Value = "Instructions:": ws.Cells(5, 1).Font.Bold = True
    ws.Cells(6, 1).Value = "1. Set tolerance in B3 above."
    ws.Cells(7, 1).Value = "2. Press Alt+F8 > RunReconciliationUltimate to re-run."
    ws.Cells(8, 1).Value = "3. Results appear in the 'Reconciliation_Report' sheet."
    
    ws.Columns("A").ColumnWidth = 35: ws.Columns("B").ColumnWidth = 15: ws.Columns("C").ColumnWidth = 50
End Sub

