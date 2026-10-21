Attribute VB_Name = "modQCLogic"
Option Explicit

' ==========================================================
' WAREHOUSE QC AUTOMATION SYSTEM - CORE BUSINESS ENGINE
' ==========================================================

Public Function GetItemSPQ(ByVal daNum As String, ByVal cleanPart As String) As Long
    GetItemSPQ = 0
    Dim wsLines As Worksheet
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    
    Dim r As Long, lastL As Long
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum Then
            If UCase(Trim(wsLines.Cells(r, "D").Value)) = UCase(Trim(cleanPart)) Then
                Dim spqCell As Variant
                spqCell = wsLines.Cells(r, "J").Value
                If IsNumeric(spqCell) And Val(spqCell) > 0 Then
                    GetItemSPQ = CLng(spqCell)
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Public Function CleanPartNumber(ByVal rawPart As String, ByVal vendorName As String) As String
    Dim s As String
    s = Trim(rawPart)
    
    If Len(s) = 0 Then
        CleanPartNumber = ""
        Exit Function
    End If
    
    Dim wsR As Worksheet, wsLines As Worksheet
    Set wsR = ThisWorkbook.Sheets("Vendor_Rules")
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    
    ' STEP 0: Check if rawPart is ALREADY an exact legitimate part number in DA lines!
    ' (Guarantees prefix-removal logic never accidentally alters a legitimate part number like PCR20-132JV)
    Dim rL_exact As Long, lastL_exact As Long
    lastL_exact = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    For rL_exact = 5 To lastL_exact
        If UCase(Trim(wsLines.Cells(rL_exact, "D").Value)) = UCase(s) Then
            CleanPartNumber = UCase(s)
            Exit Function
        End If
    Next rL_exact
    
    Dim actVendor As String
    actVendor = vendorName
    
    ' If vendor is Multi-Vendor or blank, test stripping prefixes against DA_Lines
    If UCase(Trim(actVendor)) = "MULTI-VENDOR" Or Len(Trim(actVendor)) = 0 Then
        Dim testS As String
        testS = s
        If UCase(Left(testS, 2)) = "1P" Then testS = Mid(testS, 3)
        If UCase(Left(testS, 1)) = "P" Then testS = Mid(testS, 2)
        testS = UCase(Trim(testS))
        
        Dim rL As Long, lastL As Long
        lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
        For rL = 5 To lastL
            If UCase(Trim(wsLines.Cells(rL, "D").Value)) = testS Then
                CleanPartNumber = testS
                Exit Function
            End If
        Next rL
        actVendor = "GENERIC / DEFAULT"
    End If
    
    ' Lookup rules from Vendor_Rules sheet
    Dim r As Long, lastR As Long
    lastR = wsR.Cells(wsR.Rows.Count, "B").End(xlUp).Row
    
    Dim pfxList As String, sfxTrim As Long, toUpper As Boolean
    pfxList = "1P,P"
    sfxTrim = 0
    toUpper = True
    
    For r = 5 To lastR
        If UCase(Trim(wsR.Cells(r, "B").Value)) = UCase(Trim(actVendor)) Then
            pfxList = Trim(wsR.Cells(r, "C").Value)
            sfxTrim = Val(wsR.Cells(r, "D").Value)
            toUpper = (UCase(Trim(wsR.Cells(r, "G").Value)) = "TRUE")
            Exit For
        End If
    Next r
    
    ' 1. Strip Prefixes (Case-insensitive)
    Dim pfxArray() As String, i As Long, pfx As String
    pfxArray = Split(pfxList, ",")
    For i = LBound(pfxArray) To UBound(pfxArray)
        pfx = Trim(pfxArray(i))
        If Len(pfx) > 0 Then
            If UCase(Left(s, Len(pfx))) = UCase(pfx) Then
                s = Mid(s, Len(pfx) + 1)
                Exit For
            End If
        End If
    Next i
    
    ' 2. Trim trailing suffix characters if defined (e.g. trailing batch spaces like ESR18EZPF6R20      0)
    If sfxTrim > 0 And Len(s) > sfxTrim Then
        If Right(s, 1) = " " Or Right(s, 2) = " 0" Or InStr(s, "  ") > 0 Then
            s = Left(s, Len(s) - sfxTrim)
        End If
    End If
    
    ' 3. General cleanup
    s = Trim(s)
    If toUpper Then s = UCase(s)
    
    CleanPartNumber = s
End Function

Public Function CleanQuantity(ByVal rawQty As String) As Long
    Dim s As String
    s = Trim(rawQty)
    
    If Len(s) = 0 Then
        CleanQuantity = 0
        Exit Function
    End If
    
    ' Strip Q or q prefix
    If UCase(Left(s, 1)) = "Q" Then
        s = Mid(s, 2)
    End If
    
    ' Remove commas or spaces
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    
    CleanQuantity = Val(s)
End Function

Public Sub UnprotectConsole()
    On Error Resume Next
    ThisWorkbook.Sheets("QC_Console").Unprotect
    On Error GoTo 0
End Sub

Public Sub ProtectConsole()
    On Error Resume Next
    With ThisWorkbook.Sheets("QC_Console")
        .Range("C10").Locked = False
        .Range("C12:E12").Locked = False
        .Range("G12").Locked = False
        .Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End With
    On Error GoTo 0
End Sub

Public Sub ProtectAllSheets()
    On Error Resume Next
    ProtectConsole
    ThisWorkbook.Sheets("DA_Lines").Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Daily_Registry").Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ThisWorkbook.Sheets("DA_Intake").Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Vendor_Rules").Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Printable_DA").Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Audit_Log").Protect Password:="QCAdmin2026", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub ProtectAudit()
    On Error Resume Next
    ThisWorkbook.Sheets("Audit_Log").Protect Password:="QCAdmin2026", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub WriteAuditLog(ByVal logID As Long, ByVal daNum As String, ByVal vendor As String, _
                         ByVal rawBarcode As String, ByVal cleanPart As String, ByVal qty As Long, _
                         ByVal expPart As String, ByVal expQty As Long, ByVal scnQty As Long, _
                         ByVal remQty As Long, ByVal res As String, ByVal reason As String)
    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("Audit_Log")
    On Error Resume Next
    wsAudit.Unprotect Password:="QCAdmin2026"
    On Error GoTo 0
    
    Dim lastAud As Long
    lastAud = wsAudit.Cells(wsAudit.Rows.Count, "B").End(xlUp).Row + 1
    
    wsAudit.Cells(lastAud, "B").Value = logID
    wsAudit.Cells(lastAud, "C").Value = Now
    wsAudit.Cells(lastAud, "D").Value = daNum
    wsAudit.Cells(lastAud, "E").Value = vendor
    wsAudit.Cells(lastAud, "F").Value = rawBarcode
    wsAudit.Cells(lastAud, "G").Value = cleanPart
    wsAudit.Cells(lastAud, "H").Value = qty
    wsAudit.Cells(lastAud, "I").Value = expPart
    wsAudit.Cells(lastAud, "J").Value = expQty
    wsAudit.Cells(lastAud, "K").Value = scnQty
    wsAudit.Cells(lastAud, "L").Value = remQty
    wsAudit.Cells(lastAud, "M").Value = res
    wsAudit.Cells(lastAud, "N").Value = reason
    wsAudit.Cells(lastAud, "O").Value = Environ("USERNAME")
    
    On Error Resume Next
    wsAudit.Protect Password:="QCAdmin2026", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub LoadDAConsole(Optional ByVal paramDA As String = "")
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    Dim wsQC As Worksheet, wsDB As Worksheet, wsLines As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    Set wsDB = ThisWorkbook.Sheets("Daily_Registry")
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    
    Dim daNum As String
    If Len(Trim(paramDA)) > 0 Then
        daNum = Trim(paramDA)
    Else
        daNum = Trim(wsQC.Range("C10").Value)
    End If
    
    ' Clean any scanner artifacts (* or DA prefix)
    daNum = Replace(Replace(daNum, "*", ""), " ", "")
    If UCase(Left(daNum, 2)) = "DA" Then daNum = Mid(daNum, 3)
    
    If Len(daNum) = 0 Then
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    UnprotectConsole
    
    ' Search DA_Database
    Dim r As Long, lastDB As Long, found As Boolean
    lastDB = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row
    Dim vendor As String, cust As String, daStatus As String
    
    For r = 5 To lastDB
        If CStr(wsDB.Cells(r, "B").Value) = daNum Then
            cust = wsDB.Cells(r, "D").Value
            vendor = wsDB.Cells(r, "E").Value
            daStatus = wsDB.Cells(r, "J").Value
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        wsQC.Range("B3:J4").Value = "HOLD — DA " & daNum & " NOT FOUND IN DATABASE"
        wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38) ' Red
        ProtectConsole
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    ' Update Console Inputs & KPI Cards
    wsQC.Range("C10").Value = daNum
    wsQC.Range("F10").Value = vendor
    wsQC.Range("H10").Value = cust
    wsQC.Range("H10").WrapText = True
    wsQC.Range("J10").Value = daStatus
    
    ' KPI Cards Update
    wsQC.Range("B7").Value = "DA " & daNum
    wsQC.Range("D7").Value = vendor & "  |  " & cust
    wsQC.Range("D7").WrapText = True
    
    ' Clear existing console lines (rows 16 to 45)
    wsQC.Range("B16:J45").ClearContents
    wsQC.Range("B16:J45").Interior.Color = RGB(255, 255, 255)
    
    ' Populate lines from DA_Lines
    Dim lastL As Long, outRow As Long, totExp As Long, totScn As Long, totRem As Long
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    outRow = 16
    totExp = 0: totScn = 0: totRem = 0
    
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum Then
            wsQC.Cells(outRow, "B").Value = wsLines.Cells(r, "C").Value ' Line#
            wsQC.Cells(outRow, "C").Value = wsLines.Cells(r, "D").Value ' Part#
            wsQC.Cells(outRow, "D").Value = wsLines.Cells(r, "E").Value ' Desc
            
            Dim rowVdr As String
            rowVdr = Trim(wsLines.Cells(r, "K").Value)
            If Len(rowVdr) = 0 Then rowVdr = vendor
            wsQC.Cells(outRow, "E").Value = rowVdr                      ' Vendor
            
            Dim expQ As Long, scnQ As Long, rmQ As Long
            expQ = Val(wsLines.Cells(r, "F").Value)
            scnQ = Val(wsLines.Cells(r, "G").Value)
            rmQ = Val(wsLines.Cells(r, "H").Value)
            
            totExp = totExp + expQ
            totScn = totScn + scnQ
            totRem = totRem + rmQ
            
            wsQC.Cells(outRow, "F").Value = expQ
            wsQC.Cells(outRow, "G").Value = scnQ
            wsQC.Cells(outRow, "H").Value = rmQ
            
            If expQ > 0 Then
                wsQC.Cells(outRow, "I").Value = Format(scnQ / expQ, "0%")
            Else
                wsQC.Cells(outRow, "I").Value = "0%"
            End If
            
            wsQC.Cells(outRow, "J").Value = wsLines.Cells(r, "I").Value ' Status
            
            ' Color code row
            If wsLines.Cells(r, "I").Value = "MATCHED" Then
                wsQC.Range(wsQC.Cells(outRow, "B"), wsQC.Cells(outRow, "J")).Interior.Color = RGB(220, 252, 231)
            ElseIf wsLines.Cells(r, "I").Value = "IN PROGRESS" Then
                wsQC.Range(wsQC.Cells(outRow, "B"), wsQC.Cells(outRow, "J")).Interior.Color = RGB(254, 243, 199)
            Else
                If outRow Mod 2 = 1 Then
                    wsQC.Range(wsQC.Cells(outRow, "B"), wsQC.Cells(outRow, "J")).Interior.Color = RGB(248, 250, 252)
                End If
            End If
            
            outRow = outRow + 1
        End If
    Next r
    
    ' Render dynamic Totals row
    If outRow > 16 Then
        wsQC.Cells(outRow, "D").Value = "TOTALS:"
        wsQC.Cells(outRow, "D").Font.Bold = True
        wsQC.Cells(outRow, "F").Value = totExp
        wsQC.Cells(outRow, "F").Font.Bold = True
        wsQC.Cells(outRow, "G").Value = totScn
        wsQC.Cells(outRow, "G").Font.Bold = True
        wsQC.Cells(outRow, "H").Value = totRem
        wsQC.Cells(outRow, "H").Font.Bold = True
        
        ' Ensure all QC table rows have wrap text and responsive row height
        With wsQC.Range("B16:J" & outRow)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        Dim qcR As Long
        For qcR = 16 To outRow
            If wsQC.Rows(qcR).RowHeight < 24 Then wsQC.Rows(qcR).RowHeight = 24
        Next qcR
    End If
    
    ' Update KPI Card Totals
    If totExp > 0 Then
        wsQC.Range("F7").Value = Format(totScn, "#,##0") & " / " & Format(totExp, "#,##0") & " (" & Format(totScn / totExp, "0%") & ")"
    Else
        wsQC.Range("F7").Value = "0 / 0 (0%)"
    End If
    wsQC.Range("H7").Value = Format(totRem, "#,##0") & " Pcs Remaining"
    
    ' Reset inputs
    wsQC.Range("C12").Value = ""
    wsQC.Range("G12").Value = ""
    
    ' Banner status
    If daStatus = "COMPLETE" Then
        wsQC.Range("B3:J4").Value = "DA COMPLETE — ALL ITEMS & QUANTITIES 100% VERIFIED!"
        wsQC.Range("B3:J4").Interior.Color = RGB(4, 120, 87)
    Else
        wsQC.Range("B3:J4").Value = "DA " & daNum & " LOADED (" & vendor & ") — SCAN REEL/CARTON INTO BOX [2] (SPQ AUTO-VERIFY ACTIVE)"
        wsQC.Range("B3:J4").Interior.Color = RGB(2, 132, 199)
    End If
    
    ProtectConsole
    
    ' Auto-focus component scan input
    On Error Resume Next
    wsQC.Activate
    wsQC.Range("C12").Select
    On Error GoTo 0
    
    Application.EnableEvents = prevEvents
End Sub

Public Sub ProcessScanConsole()
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    Dim wsQC As Worksheet, wsLines As Worksheet, wsAudit As Worksheet, wsDB As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    Set wsAudit = ThisWorkbook.Sheets("Audit_Log")
    Set wsDB = ThisWorkbook.Sheets("Daily_Registry")
    
    Dim daNum As String, rawBarcode As String, rawQty As String, vendor As String
    daNum = Trim(wsQC.Range("C10").Value)
    vendor = Trim(wsQC.Range("F10").Value)
    rawBarcode = Trim(wsQC.Range("C12").Value)
    rawQty = Trim(wsQC.Range("G12").Value)
    
    UnprotectConsole
    
    If Len(daNum) = 0 Then
        wsQC.Range("B3:J4").Value = "HOLD — STEP 1: SCAN OR LOAD DA NUMBER FIRST"
        wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
        ProtectConsole
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    If Len(rawBarcode) = 0 Then
        wsQC.Range("B3:J4").Value = "HOLD — STEP 2: SCAN COMPONENT REEL/CARTON BARCODE INTO BOX [2]"
        wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
        ProtectConsole
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    If Len(rawQty) = 0 Then
        wsQC.Range("B3:J4").Value = "HOLD — STEP 3: SCAN OR ENTER QUANTITY INTO BOX [3]"
        wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
        ProtectConsole
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    ' 1. Clean Part & Qty
    Dim cleanPart As String, cleanQty As Long
    cleanPart = CleanPartNumber(rawBarcode, vendor)
    cleanQty = CleanQuantity(rawQty)
    
    wsQC.Range("C12").Value = cleanPart
    
    If cleanQty <= 0 Then
        wsQC.Range("B3:J4").Value = "HOLD — INVALID QUANTITY ENTERED IN BOX [3]: " & rawQty
        wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
        ProtectConsole
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    ' 2. Deterministic exact validation against DA lines
    Dim r As Long, lastL As Long, found As Boolean, lineRow As Long
    Dim expPart As String, expQty As Long, scnQty As Long, remQty As Long, lStatus As String
    Dim completedRowFound As Long
    completedRowFound = 0
    found = False
    
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum Then
            expPart = UCase(Trim(wsLines.Cells(r, "D").Value))
            If expPart = cleanPart Then
                remQty = Val(wsLines.Cells(r, "H").Value)
                If remQty > 0 Then
                    found = True
                    lineRow = r
                    expQty = Val(wsLines.Cells(r, "F").Value)
                    scnQty = Val(wsLines.Cells(r, "G").Value)
                    lStatus = wsLines.Cells(r, "I").Value
                    Exit For
                Else
                    completedRowFound = r
                End If
            End If
        End If
    Next r
    
    Dim auditResult As String, auditReason As String
    Dim nextLogID As Long, lastAud As Long
    lastAud = wsAudit.Cells(wsAudit.Rows.Count, "B").End(xlUp).Row
    nextLogID = Val(wsAudit.Cells(lastAud, "B").Value) + 1
    If nextLogID < 1001 Then nextLogID = 1001
    
    If Not found Then
        If completedRowFound > 0 Then
            auditResult = "HOLD"
            auditReason = "LINE ALREADY COMPLETE: Requirement for " & cleanPart & " is already fully satisfied"
            wsQC.Range("B3:J4").Value = "HOLD — LINE ALREADY COMPLETED (" & cleanPart & ")"
            wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
            expQty = Val(wsLines.Cells(completedRowFound, "F").Value)
            scnQty = Val(wsLines.Cells(completedRowFound, "G").Value)
            WriteAuditLog nextLogID, daNum, vendor, rawBarcode, cleanPart, cleanQty, cleanPart, expQty, scnQty, 0, auditResult, auditReason
        Else
            auditResult = "HOLD"
            auditReason = "PART NOT ON DA: Scanned part " & cleanPart & " does not match DA " & daNum
            wsQC.Range("B3:J4").Value = "HOLD — PART NOT ON DA (" & cleanPart & ")"
            wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
            WriteAuditLog nextLogID, daNum, vendor, rawBarcode, cleanPart, cleanQty, "N/A", 0, 0, 0, auditResult, auditReason
        End If
        ProtectConsole
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    ' Check if quantity exceeds DA remaining
    If cleanQty > remQty Then
        auditResult = "HOLD"
        auditReason = "QUANTITY EXCEEDS DA: Entered " & cleanQty & " exceeds remaining requirement " & remQty
        wsQC.Range("B3:J4").Value = "HOLD — QUANTITY EXCEEDS DA (Entered: " & cleanQty & " | Remaining: " & remQty & ")"
        wsQC.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
        
        WriteAuditLog nextLogID, daNum, vendor, rawBarcode, cleanPart, cleanQty, expPart, expQty, scnQty, remQty, auditResult, auditReason
        ProtectConsole
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    ' VALID MATCH! Update cumulative totals
    wsLines.Unprotect
    wsDB.Unprotect
    
    Dim newScn As Long, newRem As Long, newStatus As String
    newScn = scnQty + cleanQty
    newRem = expQty - newScn
    If newRem = 0 Then
        newStatus = "MATCHED"
    Else
        newStatus = "IN PROGRESS"
    End If
    
    wsLines.Cells(lineRow, "G").Value = newScn
    wsLines.Cells(lineRow, "H").Value = newRem
    wsLines.Cells(lineRow, "I").Value = newStatus
    
    ' Check if SPQ was auto-applied
    Dim spqUsedMsg As String
    Dim lineSPQ As Long
    lineSPQ = Val(wsLines.Cells(lineRow, "J").Value)
    If lineSPQ > 0 And cleanQty = lineSPQ Then
        spqUsedMsg = " [SPQ AUTO-VERIFIED: " & Format(cleanQty, "#,##0") & " Pcs]"
    Else
        spqUsedMsg = ""
    End If
    
    ' Log audit
    auditResult = "MATCH"
    auditReason = "PASS: " & cleanPart & " +" & cleanQty & " pcs" & spqUsedMsg & " (Checked: " & newScn & "/" & expQty & ")"
    WriteAuditLog nextLogID, daNum, vendor, rawBarcode, cleanPart, cleanQty, expPart, expQty, newScn, newRem, auditResult, auditReason
    
    ' Check if all lines for DA are complete
    Dim allComplete As Boolean: allComplete = True
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum Then
            If Val(wsLines.Cells(r, "H").Value) > 0 Then
                allComplete = False
                Exit For
            End If
        End If
    Next r
    
    ' Update Daily_Registry (Col G=Exp, Col H=Scn, Col I=Rem, Col J=Status)
    Dim lastDB As Long
    lastDB = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row
    For r = 5 To lastDB
        If CStr(wsDB.Cells(r, "B").Value) = daNum Then
            Dim curDBScn As Long
            curDBScn = Val(wsDB.Cells(r, "H").Value) + cleanQty
            wsDB.Cells(r, "H").Value = curDBScn
            wsDB.Cells(r, "I").Value = Val(wsDB.Cells(r, "G").Value) - curDBScn
            If allComplete Then
                wsDB.Cells(r, "J").Value = "COMPLETE"
            End If
            Exit For
        End If
    Next r
    
    wsLines.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    wsDB.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    ' Refresh Console Display
    LoadDAConsole daNum
    
    UnprotectConsole
    If allComplete Then
        wsQC.Range("J10").Value = "COMPLETE"
        wsQC.Range("B3:J4").Value = "DA COMPLETE — ALL ITEMS & QUANTITIES 100% VERIFIED!"
        wsQC.Range("B3:J4").Interior.Color = RGB(4, 120, 87)
    Else
        wsQC.Range("B3:J4").Value = "MATCH — PROCEED! [ Verified " & cleanPart & " (+" & Format(cleanQty, "#,##0") & " Pcs" & IIf(Len(spqUsedMsg) > 0, " SPQ Auto-Applied", "") & ") ] — SCAN NEXT CARTON"
        wsQC.Range("B3:J4").Interior.Color = RGB(5, 150, 105)
    End If
    
    ' Reset scan inputs and refocus Box [2] for rapid-fire carton scanning
    wsQC.Range("C12").Value = ""
    wsQC.Range("G12").Value = ""
    ProtectConsole
    
    On Error Resume Next
    wsQC.Activate
    wsQC.Range("C12").Select
    On Error GoTo 0
    
    Application.EnableEvents = prevEvents
End Sub

Public Sub ResetConsole()
    UnprotectConsole
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    wsQC.Range("C12").Value = ""
    wsQC.Range("G12").Value = ""
    wsQC.Range("B3:J4").Value = "READY — SCAN COMPONENT BARCODE INTO BOX [2]"
    wsQC.Range("B3:J4").Interior.Color = RGB(2, 132, 199)
    ProtectConsole
    On Error Resume Next
    wsQC.Activate
    wsQC.Range("C12").Select
    On Error GoTo 0
End Sub

Public Sub ExecuteDeleteCurrentScan()
    UnprotectConsole
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    
    Dim prevScan As String
    prevScan = Trim(wsQC.Range("C12").Value)
    
    wsQC.Range("C12").Value = ""
    wsQC.Range("G12").Value = ""
    
    Dim daNum As String
    daNum = Trim(wsQC.Range("C10").Value)
    
    wsQC.Range("B3:J4").Value = "SCAN DELETED — SCANNER READY FOR NEXT REEL/CARTON BARCODE"
    wsQC.Range("B3:J4").Interior.Color = RGB(2, 132, 199) ' Sky Blue
    
    ProtectConsole
    
    On Error Resume Next
    wsQC.Activate
    wsQC.Range("C12").Select
    On Error GoTo 0
    
    If Application.Visible And Len(prevScan) > 0 Then
        Application.StatusBar = "Scanned barcode [" & prevScan & "] deleted. Ready for next scan."
    End If
End Sub

Public Sub GetLastVerifiedScan(ByVal daNum As String, ByRef outPart As String, ByRef outQty As Long)
    outPart = ""
    outQty = 0
    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("Audit_Log")
    
    Dim r As Long, lastAud As Long
    lastAud = wsAudit.Cells(wsAudit.Rows.Count, "B").End(xlUp).Row
    
    For r = lastAud To 5 Step -1
        If CStr(wsAudit.Cells(r, "D").Value) = daNum And wsAudit.Cells(r, "M").Value = "MATCH" Then
            outPart = Trim(wsAudit.Cells(r, "G").Value)
            outQty = Val(wsAudit.Cells(r, "H").Value)
            Exit For
        End If
    Next r
End Sub

Public Sub ExecuteUndoLastScan(ByVal daNum As String, ByVal targetPart As String, ByVal revQty As Long)
    If Len(daNum) = 0 Or Len(targetPart) = 0 Or revQty <= 0 Then Exit Sub
    
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    Dim wsLines As Worksheet, wsDB As Worksheet, wsAudit As Worksheet, wsQC As Worksheet
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    Set wsDB = ThisWorkbook.Sheets("Daily_Registry")
    Set wsAudit = ThisWorkbook.Sheets("Audit_Log")
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    
    ' 1. Find line in DA_Lines and reduce scanned qty
    wsLines.Unprotect
    wsDB.Unprotect
    
    Dim r As Long, lastL As Long, found As Boolean
    Dim curScn As Long, curExp As Long, newScn As Long, newRem As Long, newStatus As String
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum And UCase(Trim(wsLines.Cells(r, "D").Value)) = UCase(targetPart) Then
            curExp = Val(wsLines.Cells(r, "F").Value)
            curScn = Val(wsLines.Cells(r, "G").Value)
            newScn = curScn - revQty
            If newScn < 0 Then newScn = 0
            newRem = curExp - newScn
            If newScn = 0 Then
                newStatus = "PENDING"
            ElseIf newRem > 0 Then
                newStatus = "IN PROGRESS"
            Else
                newStatus = "MATCHED"
            End If
            
            wsLines.Cells(r, "G").Value = newScn
            wsLines.Cells(r, "H").Value = newRem
            wsLines.Cells(r, "I").Value = newStatus
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        wsLines.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        wsDB.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        If Application.Visible Then
            MsgBox "Could not locate line item " & targetPart & " on DA " & daNum & ".", vbExclamation, "Rollback Error"
        End If
        Application.EnableEvents = prevEvents
        Exit Sub
    End If
    
    ' 2. Update Daily_Registry (Col G=Exp, Col H=Scn, Col I=Rem, Col J=Status)
    Dim lastDB As Long
    lastDB = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row
    For r = 5 To lastDB
        If CStr(wsDB.Cells(r, "B").Value) = daNum Then
            Dim dbScn As Long
            dbScn = Val(wsDB.Cells(r, "H").Value) - revQty
            If dbScn < 0 Then dbScn = 0
            wsDB.Cells(r, "H").Value = dbScn
            wsDB.Cells(r, "I").Value = Val(wsDB.Cells(r, "G").Value) - dbScn
            wsDB.Cells(r, "J").Value = "OPEN"
            Exit For
        End If
    Next r
    
    wsLines.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    wsDB.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    ' 3. Log to Audit_Log
    Dim lastAud As Long, nextLogID As Long
    lastAud = wsAudit.Cells(wsAudit.Rows.Count, "B").End(xlUp).Row
    nextLogID = Val(wsAudit.Cells(lastAud, "B").Value) + 1
    If nextLogID < 1001 Then nextLogID = 1001
    
    WriteAuditLog nextLogID, daNum, Trim(wsQC.Range("F10").Value), "(OPERATOR_DELETE)", targetPart, revQty, targetPart, curExp, newScn, newRem, "REVERT", "OPERATOR DELETE: Cancelled scan of " & targetPart & " (-" & revQty & " Pcs)"
    
    ' 4. Refresh Console Display
    LoadDAConsole daNum
    
    UnprotectConsole
    wsQC.Range("B3:J4").Value = "REVERTED — Scan of " & targetPart & " (-" & Format(revQty, "#,##0") & " Pcs) cancelled."
    wsQC.Range("B3:J4").Interior.Color = RGB(234, 88, 12) ' Warm Orange
    wsQC.Range("C12").Value = ""
    wsQC.Range("G12").Value = ""
    ProtectConsole
    
    On Error Resume Next
    wsQC.Activate
    wsQC.Range("C12").Select
    On Error GoTo 0
    
    If Application.Visible Then
        MsgBox "Scan successfully rolled back:" & vbCrLf & vbCrLf & _
               "• Component: " & targetPart & vbCrLf & _
               "• Reverted Qty: -" & Format(revQty, "#,##0") & " Pcs" & vbCrLf & _
               "• Remaining Requirement: " & Format(newRem, "#,##0") & " Pcs" & vbCrLf & _
               "• Status: " & newStatus, vbInformation, "Scan Rolled Back"
    End If
    
    Application.EnableEvents = prevEvents
End Sub

Public Sub DeleteOrClearScan()
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    
    Dim rawScan As String, cleanPart As String, rawQty As String, daNum As String
    daNum = Trim(wsQC.Range("C10").Value)
    rawScan = Trim(wsQC.Range("C12").Value)
    cleanPart = CleanPartNumber(rawScan, Trim(wsQC.Range("F10").Value))
    rawQty = Trim(wsQC.Range("G12").Value)
    
    ' Try to show modal popout form
    On Error Resume Next
    frmConfirmDelete.Show vbModal
    If Err.Number = 0 Then Exit Sub
    On Error GoTo 0
    
    ' Native MsgBox Fallback if form is unavailable
    If Len(rawScan) > 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("CONFIRM DELETE SCANNED BARCODE:" & vbCrLf & vbCrLf & _
                      "• Scanned Barcode: " & rawScan & vbCrLf & _
                      "• Cleaned Part No: " & cleanPart & vbCrLf & _
                      IIf(Len(rawQty) > 0, "• Quantity: " & rawQty & " Pcs" & vbCrLf, "") & vbCrLf & _
                      "Are you sure you want to delete and clear this barcode from the scanner?", _
                      vbYesNo + vbExclamation, "Delete Scanned Barcode")
        If resp = vbYes Then
            ExecuteDeleteCurrentScan
        End If
    Else
        Dim lastP As String, lastQ As Long
        GetLastVerifiedScan daNum, lastP, lastQ
        If Len(lastP) > 0 And lastQ > 0 Then
            Dim respUndo As VbMsgBoxResult
            respUndo = MsgBox("Scan box [2] is empty." & vbCrLf & vbCrLf & _
                              "Recent Verified Scan Found:" & vbCrLf & _
                              "• Part Number: " & lastP & vbCrLf & _
                              "• Quantity: " & Format(lastQ, "#,##0") & " Pcs" & vbCrLf & vbCrLf & _
                              "Do you want to UNDO / DELETE this scan and restore the DA remaining requirement?", _
                              vbYesNo + vbQuestion, "Undo Recent Scan")
            If respUndo = vbYes Then
                ExecuteUndoLastScan daNum, lastP, lastQ
            End If
        Else
            MsgBox "There is no barcode in scan box [2], and no recent scans to undo.", vbInformation, "Scan Box Empty"
        End If
    End If
End Sub

Public Sub ShowScannerForm()
    On Error Resume Next
    frmQCScanner.Show vbModeless
    On Error GoTo 0
End Sub

Public Sub ShowSPQForm()
    On Error Resume Next
    frmSPQInput.Show vbModal
    On Error GoTo 0
End Sub

Public Sub NavigateDailyRegistry()
    On Error Resume Next
    ThisWorkbook.Sheets("Daily_Registry").Activate
    On Error GoTo 0
End Sub

Public Sub NavigateVendorRules()
    On Error Resume Next
    ThisWorkbook.Sheets("Vendor_Rules").Activate
    On Error GoTo 0
End Sub

Public Sub NavigateQCConsole()
    On Error Resume Next
    ThisWorkbook.Sheets("QC_Console").Activate
    ThisWorkbook.Sheets("QC_Console").Range("C12").Select
    On Error GoTo 0
End Sub

Public Sub NavigateDAIntake()
    On Error Resume Next
    ThisWorkbook.Sheets("DA_Intake").Activate
    On Error GoTo 0
End Sub

Public Function IsPartOnDA(ByVal daNum As String, ByVal cleanPart As String) As Boolean
    IsPartOnDA = False
    Dim wsLines As Worksheet
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    Dim r As Long, lastL As Long
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum And UCase(Trim(wsLines.Cells(r, "D").Value)) = UCase(cleanPart) Then
            IsPartOnDA = True
            Exit Function
        End If
    Next r
End Function

Public Sub PromptQuantityForPart(ByVal partNo As String)
    On Error Resume Next
    frmQuantityInput.SetupPart partNo
    frmQuantityInput.Show vbModal
    On Error GoTo 0
End Sub


