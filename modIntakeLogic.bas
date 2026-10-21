Attribute VB_Name = "modIntakeLogic"
Option Explicit

Private Function CleanNumericStr(ByVal s As String) As Long
    Dim i As Long, ch As String, numStr As String
    numStr = ""
    s = Replace(Replace(s, ",", ""), " ", "")
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
        ElseIf Len(numStr) > 0 Then
            Exit For
        End If
    Next i
    If Len(numStr) > 0 Then CleanNumericStr = Val(numStr) Else CleanNumericStr = 0
End Function

Public Sub ParseDAEmailText()
    On Error GoTo ErrHandler
    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Sheets("DA_Intake")
    
    Dim rawText As String
    rawText = Trim(CStr(wsIn.Range("B7").Value))
    
    ' Check if text might be in other intake cells
    If Len(rawText) = 0 Then
        Dim c As Range
        For Each c In wsIn.Range("B7:H13")
            If Len(Trim(CStr(c.Value))) > 0 Then
                rawText = rawText & vbLf & Trim(CStr(c.Value))
            End If
        Next c
    End If
    
    If Len(Trim(rawText)) = 0 Then
        If Application.Visible Then
            MsgBox "Please paste a DA email or packing slip into Cell B7 first." & vbCrLf & vbCrLf & _
                   "Tip: Select Cell B7, press Ctrl+V to paste your email text, then click [ 1-Click Parse Email ].", _
                   vbExclamation, "No Text Found"
        End If
        Exit Sub
    End If
    
    ' Normalize line breaks
    rawText = Replace(rawText, vbCrLf, vbLf)
    rawText = Replace(rawText, vbCr, vbLf)
    
    Dim lines() As String, i As Long, lineStr As String
    lines = Split(rawText, vbLf)
    
    Dim daNum As String, vendor As String, cust As String
    daNum = "": vendor = "": cust = ""
    
    ' 1. Parse Header Fields
    For i = LBound(lines) To UBound(lines)
        lineStr = Trim(lines(i))
        If Len(lineStr) > 0 Then
            Dim upLine As String
            upLine = UCase(lineStr)
            
            ' DA / CRM Dispatch Number / Packing List Number
            If Len(daNum) = 0 Then
                Dim keyPos As Long, keyLen As Long
                keyPos = 0: keyLen = 0
                
                If InStr(upLine, "CRM DISPATCH NUMBER") > 0 Then
                    keyPos = InStr(upLine, "CRM DISPATCH NUMBER"): keyLen = 19
                ElseIf InStr(upLine, "DISPATCH NUMBER") > 0 Then
                    keyPos = InStr(upLine, "DISPATCH NUMBER"): keyLen = 15
                ElseIf InStr(upLine, "PACKING LIST NO:") > 0 Then
                    keyPos = InStr(upLine, "PACKING LIST NO:"): keyLen = 16
                ElseIf InStr(upLine, "PACKING LIST NO.") > 0 Then
                    keyPos = InStr(upLine, "PACKING LIST NO."): keyLen = 16
                ElseIf InStr(upLine, "PACKING LIST:") > 0 Then
                    keyPos = InStr(upLine, "PACKING LIST:"): keyLen = 13
                ElseIf InStr(upLine, "DISPATCH ADVICE NO:") > 0 Then
                    keyPos = InStr(upLine, "DISPATCH ADVICE NO:"): keyLen = 19
                ElseIf InStr(upLine, "DISPATCH ADVICE #") > 0 Then
                    keyPos = InStr(upLine, "DISPATCH ADVICE #"): keyLen = 17
                ElseIf InStr(upLine, "DA NUMBER:") > 0 Then
                    keyPos = InStr(upLine, "DA NUMBER:"): keyLen = 10
                ElseIf InStr(upLine, "DA NO:") > 0 Then
                    keyPos = InStr(upLine, "DA NO:"): keyLen = 6
                ElseIf InStr(upLine, "DA #:") > 0 Then
                    keyPos = InStr(upLine, "DA #:"): keyLen = 5
                ElseIf InStr(upLine, "DA#:") > 0 Then
                    keyPos = InStr(upLine, "DA#:"): keyLen = 4
                ElseIf InStr(upLine, "DA#") > 0 Then
                    keyPos = InStr(upLine, "DA#"): keyLen = 3
                ElseIf InStr(upLine, "DA:") > 0 Then
                    keyPos = InStr(upLine, "DA:"): keyLen = 3
                ElseIf InStr(upLine, "PL NO:") > 0 Then
                    keyPos = InStr(upLine, "PL NO:"): keyLen = 6
                End If
                
                If keyPos > 0 Then
                    Dim rawDA As String
                    rawDA = Trim(Mid(lineStr, keyPos + keyLen))
                    Dim spPos As Long
                    spPos = InStr(rawDA, " ")
                    If spPos > 0 Then
                        Dim tbPos As Long
                        tbPos = InStr(rawDA, vbTab)
                        If tbPos > 0 And tbPos < spPos Then spPos = tbPos
                        daNum = Trim(Left(rawDA, spPos - 1))
                    Else
                        daNum = rawDA
                    End If
                    daNum = Replace(Replace(Replace(Replace(Replace(Replace(Replace(daNum, ":", ""), "#", ""), "*", ""), "-", ""), vbTab, ""), ",", ""), ".", "")
                    daNum = Trim(daNum)
                End If
            End If
            
            ' Consignee / Customer
            If Len(cust) = 0 Then
                If InStr(upLine, " FOR ") > 0 And InStr(upLine, " AGAINST") > 0 Then
                    Dim pFor As Long, pAgn As Long
                    pFor = InStr(upLine, " FOR ") + 5
                    pAgn = InStr(upLine, " AGAINST")
                    If pAgn > pFor Then
                        cust = Trim(Mid(lineStr, pFor, pAgn - pFor))
                        If InStr(cust, ",") > 0 Then cust = Trim(Left(cust, InStr(cust, ",") - 1))
                    End If
                ElseIf InStr(upLine, "CONSIGNEE:") > 0 Then
                    cust = Trim(Mid(lineStr, InStr(upLine, "CONSIGNEE:") + 10))
                ElseIf InStr(upLine, "CUSTOMER:") > 0 Then
                    cust = Trim(Mid(lineStr, InStr(upLine, "CUSTOMER:") + 9))
                ElseIf InStr(upLine, "SHIP TO:") > 0 Then
                    cust = Trim(Mid(lineStr, InStr(upLine, "SHIP TO:") + 8))
                ElseIf InStr(upLine, "BUYER:") > 0 Then
                    cust = Trim(Mid(lineStr, InStr(upLine, "BUYER:") + 6))
                End If
                If InStr(cust, "|") > 0 Then cust = Left(cust, InStr(cust, "|") - 1)
                If InStr(cust, vbTab) > 0 Then cust = Left(cust, InStr(cust, vbTab) - 1)
                cust = Trim(cust)
            End If
            
            ' Vendor / Make
            If Len(vendor) = 0 Then
                If InStr(upLine, "MAKE:") > 0 And InStr(lineStr, "|") = 0 Then
                    vendor = Trim(Mid(lineStr, InStr(upLine, "MAKE:") + 5))
                ElseIf InStr(upLine, "VENDOR:") > 0 And InStr(lineStr, "|") = 0 Then
                    vendor = Trim(Mid(lineStr, InStr(upLine, "VENDOR:") + 7))
                ElseIf InStr(upLine, "SUPPLIER:") > 0 And InStr(lineStr, "|") = 0 Then
                    vendor = Trim(Mid(lineStr, InStr(upLine, "SUPPLIER:") + 9))
                ElseIf InStr(upLine, "MANUFACTURER:") > 0 And InStr(lineStr, "|") = 0 Then
                    vendor = Trim(Mid(lineStr, InStr(upLine, "MANUFACTURER:") + 13))
                End If
                If InStr(vendor, "|") > 0 Then vendor = Left(vendor, InStr(vendor, "|") - 1)
                If InStr(vendor, vbTab) > 0 Then vendor = Left(vendor, InStr(vendor, vbTab) - 1)
                vendor = Trim(vendor)
            End If
        End If
    Next i
    
    ' Populate header cells if detected
    If Len(daNum) > 0 Then wsIn.Range("C4").Value = daNum
    If Len(cust) > 0 Then
        wsIn.Range("G4").Value = cust
        wsIn.Range("G4:H4").WrapText = True
        wsIn.Range("G4:H4").VerticalAlignment = xlCenter
        wsIn.Range("G4:H4").Font.Size = 10
    End If
    
    ' 2. Parse Line Items into table rows 17+
    wsIn.Range("B17:H60").ClearContents
    
    Dim outR As Long, itemCount As Long, totalQty As Long
    outR = 17
    itemCount = 0
    totalQty = 0
    
    ' Column index mapping for tabular tables (0-indexed)
    Dim colItemIdx As Long, colMakeIdx As Long, colDescIdx As Long, colQtyIdx As Long, colSPQIdx As Long
    colItemIdx = -1: colMakeIdx = -1: colDescIdx = -1: colQtyIdx = -1: colSPQIdx = -1
    
    Dim vendorsFound As String
    vendorsFound = ""
    
    For i = LBound(lines) To UBound(lines)
        lineStr = Trim(lines(i))
        If Len(lineStr) > 0 Then
            Dim upL As String
            upL = UCase(lineStr)
            
            ' Ignore header summary / title lines
            If InStr(upL, "TOTAL SHIPMENT QUANTITY") = 0 And _
               InStr(upL, "TOTAL QUANTITY") = 0 And _
               InStr(upL, "TOTAL:") = 0 And _
               InStr(upL, "PACKING LIST") = 0 And _
               InStr(upL, "DISPATCH ADVICE") = 0 And _
               InStr(upL, "CONSIGNEE:") = 0 And _
               InStr(upL, "SHIPMENT ENTRY NEEDS TO BE CREATED") = 0 And _
               InStr(upL, "DISPATCH ITEM DETAILS:") = 0 And _
               InStr(upL, "STATUS: READY") = 0 Then
               
                ' Check if this line is the table header row
                If (InStr(upL, "ITEM NO") > 0 Or InStr(upL, "ITEM NO.") > 0 Or InStr(upL, "CPN") > 0 Or InStr(upL, "SO NO") > 0) And _
                   (InStr(upL, "QTY") > 0 Or InStr(upL, "DISPATCH") > 0) Then
                    
                    Dim hParts() As String, hIdx As Long, hText As String
                    If InStr(lineStr, vbTab) > 0 Then
                        hParts = Split(lineStr, vbTab)
                    ElseIf InStr(lineStr, "|") > 0 Then
                        hParts = Split(lineStr, "|")
                    Else
                        ' Space separated headers - normalize spaces
                        Dim normH As String: normH = lineStr
                        Do While InStr(normH, "  ") > 0
                            normH = Replace(normH, "  ", " ")
                        Loop
                        normH = Replace(Replace(normH, "Sr.No.", "SrNo"), "Sr. No.", "SrNo")
                        normH = Replace(normH, "SO No", "SONo")
                        normH = Replace(Replace(normH, "Item No.", "ItemNo"), "Item No", "ItemNo")
                        normH = Replace(normH, "Dispatch QTY", "DispatchQty")
                        normH = Replace(normH, "Strategic Item", "StrategicItem")
                        hParts = Split(normH, " ")
                    End If
                    
                    For hIdx = LBound(hParts) To UBound(hParts)
                        hText = UCase(Trim(hParts(hIdx)))
                        If InStr(hText, "ITEMNO") > 0 Or InStr(hText, "ITEM NO") > 0 Or InStr(hText, "ITEM NO.") > 0 Or InStr(hText, "PART NO") > 0 Then
                            If InStr(hText, "SR") = 0 And InStr(hText, "SO") = 0 Then colItemIdx = hIdx
                        ElseIf InStr(hText, "MAKE") > 0 Or InStr(hText, "VENDOR") > 0 Or InStr(hText, "MFR") > 0 Then
                            colMakeIdx = hIdx
                        ElseIf InStr(hText, "DESCRIPTION") > 0 Or InStr(hText, "DESC") > 0 Then
                            colDescIdx = hIdx
                        ElseIf InStr(hText, "DISPATCHQTY") > 0 Or InStr(hText, "DISPATCH QTY") > 0 Or (InStr(hText, "QTY") > 0 And InStr(hText, "SPQ") = 0) Then
                            If colQtyIdx = -1 Then colQtyIdx = hIdx
                        ElseIf InStr(hText, "SPQ") > 0 Then
                            colSPQIdx = hIdx
                        End If
                    Next hIdx
                Else
                    ' Data row parsing
                    Dim pPart As String, pDesc As String, pVdr As String, pQty As Long, pSPQ As Long
                    pPart = "": pDesc = "": pVdr = vendor: pQty = 0: pSPQ = 0
                    
                    Dim hasDelim As Boolean
                    hasDelim = False
                    
                    Dim parts() As String
                    If InStr(lineStr, vbTab) > 0 Then
                        parts = Split(lineStr, vbTab)
                        hasDelim = True
                    ElseIf InStr(lineStr, "|") > 0 Then
                        parts = Split(lineStr, "|")
                        hasDelim = True
                    ElseIf InStr(lineStr, ";") > 0 Then
                        parts = Split(lineStr, ";")
                        hasDelim = True
                    ElseIf InStr(lineStr, ",") > 0 And (InStr(upL, "QTY") > 0 Or InStr(upL, "MAKE") > 0) Then
                        parts = Split(lineStr, ",")
                        hasDelim = True
                    End If
                    
                    If hasDelim Then
                        If UBound(parts) >= 1 Then
                            ' If header was not explicitly found, set default 10-column SAP mapping if line has >= 7 parts and numeric first element
                            If colItemIdx = -1 And UBound(parts) >= 6 Then
                                If IsNumeric(Trim(parts(0))) Then
                                    colItemIdx = 2
                                    colMakeIdx = 3
                                    colDescIdx = 4
                                    colQtyIdx = 6
                                    If UBound(parts) >= 7 Then colSPQIdx = 7
                                End If
                            End If
                            
                            If colItemIdx >= 0 And colItemIdx <= UBound(parts) Then
                                pPart = Trim(parts(colItemIdx))
                                If colDescIdx >= 0 And colDescIdx <= UBound(parts) Then pDesc = Trim(parts(colDescIdx))
                                If colMakeIdx >= 0 And colMakeIdx <= UBound(parts) Then pVdr = Trim(parts(colMakeIdx))
                                If colQtyIdx >= 0 And colQtyIdx <= UBound(parts) Then pQty = CleanNumericStr(parts(colQtyIdx))
                                If colSPQIdx >= 0 And colSPQIdx <= UBound(parts) Then pSPQ = CleanNumericStr(parts(colSPQIdx))
                            Else
                                ' Fallback for 2-4 segment formats (e.g. Item 1 | Part | Desc | Qty)
                                Dim seg0 As String, seg1 As String
                                seg0 = Trim(parts(0))
                                seg1 = Trim(parts(1))
                                
                                If (InStr(UCase(seg0), "ITEM") > 0 And Len(seg0) <= 8) Or IsNumeric(seg0) Then
                                    pPart = seg1
                                    If UBound(parts) >= 2 Then pDesc = Trim(parts(2))
                                Else
                                    If InStr(seg0, ":") > 0 Then seg0 = Trim(Mid(seg0, InStr(seg0, ":") + 1))
                                    pPart = seg0
                                    pDesc = seg1
                                End If
                                
                                Dim j As Long, sSeg As String, upSeg As String
                                For j = 1 To UBound(parts)
                                    sSeg = Trim(parts(j))
                                    upSeg = UCase(sSeg)
                                    If InStr(upSeg, "MAKE:") > 0 Then
                                        pVdr = Trim(Mid(sSeg, InStr(upSeg, "MAKE:") + 5))
                                    ElseIf InStr(upSeg, "VENDOR:") > 0 Then
                                        pVdr = Trim(Mid(sSeg, InStr(upSeg, "VENDOR:") + 7))
                                    End If
                                    If InStr(upSeg, "SPQ:") > 0 Then
                                        pSPQ = CleanNumericStr(Mid(sSeg, InStr(upSeg, "SPQ:") + 4))
                                    ElseIf InStr(upSeg, "QTY") > 0 Then
                                        pQty = CleanNumericStr(sSeg)
                                    ElseIf pQty = 0 Then
                                        Dim cleanVal As Long
                                        cleanVal = CleanNumericStr(sSeg)
                                        If cleanVal > 0 Then pQty = cleanVal
                                    End If
                                Next j
                            End If
                        End If
                    Else
                        ' Space-delimited table line fallback (e.g. plain text paste)
                        Dim sTokens() As String
                        sTokens = Split(lineStr, " ")
                        Dim cleanTokens() As String, tCount As Long, tIdx As Long
                        tCount = 0
                        ReDim cleanTokens(0 To UBound(sTokens))
                        For tIdx = LBound(sTokens) To UBound(sTokens)
                            If Len(Trim(sTokens(tIdx))) > 0 Then
                                cleanTokens(tCount) = Trim(sTokens(tIdx))
                                tCount = tCount + 1
                            End If
                        Next tIdx
                        
                        If tCount >= 8 Then
                            If IsNumeric(cleanTokens(0)) And IsNumeric(cleanTokens(1)) Then
                                pPart = cleanTokens(2)
                                pVdr = cleanTokens(3)
                                Dim endDescIdx As Long
                                If UCase(cleanTokens(tCount - 1)) = "N" Or UCase(cleanTokens(tCount - 1)) = "Y" Then
                                    ' Format ends with: CPN [t-5], QTY [t-4], SPQ [t-3], Location [t-2], Strategic Item [t-1]
                                    pQty = CleanNumericStr(cleanTokens(tCount - 4))
                                    pSPQ = CleanNumericStr(cleanTokens(tCount - 3))
                                    endDescIdx = tCount - 6
                                ElseIf IsNumeric(cleanTokens(tCount - 1)) And IsNumeric(cleanTokens(tCount - 2)) Then
                                    ' Format ends with: CPN [t-3], QTY [t-2], SPQ [t-1]
                                    pQty = CleanNumericStr(cleanTokens(tCount - 2))
                                    pSPQ = CleanNumericStr(cleanTokens(tCount - 1))
                                    endDescIdx = tCount - 4
                                Else
                                    pQty = CleanNumericStr(cleanTokens(tCount - 4))
                                    pSPQ = CleanNumericStr(cleanTokens(tCount - 3))
                                    endDescIdx = tCount - 6
                                End If
                                If endDescIdx < 4 Then endDescIdx = 4
                                pDesc = ""
                                Dim dIdx As Long
                                For dIdx = 4 To endDescIdx
                                    pDesc = pDesc & cleanTokens(dIdx) & " "
                                Next dIdx
                                pDesc = Trim(pDesc)
                            End If
                        End If
                    End If
                    
                    ' Clean up part number
                    If InStr(UCase(pPart), "PART NO:") > 0 Then pPart = Trim(Mid(pPart, InStr(UCase(pPart), "PART NO:") + 8))
                    If InStr(UCase(pPart), "PART NO.") > 0 Then pPart = Trim(Mid(pPart, InStr(UCase(pPart), "PART NO.") + 8))
                    If InStr(UCase(pPart), "P/N:") > 0 Then pPart = Trim(Mid(pPart, InStr(UCase(pPart), "P/N:") + 4))
                    pPart = UCase(Trim(pPart))
                    pPart = Replace(Replace(Replace(pPart, "*", ""), "#", ""), vbTab, "")
                    
                    If Len(pVdr) = 0 Then pVdr = vendor
                    
                    If Len(pPart) > 0 And pQty > 0 And UCase(pPart) <> "ITEM NO." And UCase(pPart) <> "ITEM NO" Then
                        itemCount = itemCount + 1
                        totalQty = totalQty + pQty
                        If pSPQ <= 0 Then pSPQ = 0
                        
                        wsIn.Cells(outR, "B").Value = itemCount
                        wsIn.Cells(outR, "C").Value = pPart
                        wsIn.Cells(outR, "D").Value = pDesc
                        wsIn.Cells(outR, "E").Value = pVdr
                        wsIn.Cells(outR, "F").Value = pQty
                        wsIn.Cells(outR, "G").Value = pSPQ
                        wsIn.Cells(outR, "H").Value = "Pcs"
                        
                        If Len(pVdr) > 0 Then
                            If InStr(vendorsFound, "|" & UCase(pVdr) & "|") = 0 Then
                                vendorsFound = vendorsFound & "|" & UCase(pVdr) & "|"
                            End If
                        End If
                        
                        outR = outR + 1
                    End If
                End If
            End If
        End If
    Next i
    
    ' Apply WrapText and responsive row heights to parsed line items
    If outR > 17 Then
        Dim maxInR As Long: maxInR = outR - 1
        With wsIn.Range("B17:H" & maxInR)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        Dim rInAdj As Long
        For rInAdj = 17 To maxInR
            If wsIn.Rows(rInAdj).RowHeight < 24 Then wsIn.Rows(rInAdj).RowHeight = 24
        Next rInAdj
    End If
    
    ' Determine Vendor name
    If Len(vendor) = 0 Then
        Dim vList() As String
        vList = Split(vendorsFound, "|")
        Dim vCount As Long, k As Long, singleVdr As String
        vCount = 0: singleVdr = ""
        For k = LBound(vList) To UBound(vList)
            If Len(Trim(vList(k))) > 0 Then
                vCount = vCount + 1
                singleVdr = Trim(vList(k))
            End If
        Next k
        If vCount = 1 Then
            vendor = singleVdr
        ElseIf vCount > 1 Then
            vendor = "Multi-Vendor"
        Else
            vendor = "Multi-Vendor"
        End If
    End If
    
    If Len(vendor) > 0 Then wsIn.Range("E4").Value = vendor
    
    If Application.Visible Then
        If itemCount > 0 Then
            MsgBox "DA Email successfully parsed!" & vbCrLf & vbCrLf & _
                   "• DA Number: " & wsIn.Range("C4").Value & vbCrLf & _
                   "• Vendor: " & wsIn.Range("E4").Value & vbCrLf & _
                   "• Consignee: " & wsIn.Range("G4").Value & vbCrLf & _
                   "• Line Items Parsed: " & itemCount & " items" & vbCrLf & _
                   "• Total Quantity: " & Format(totalQty, "#,##0") & " Pcs" & vbCrLf & _
                   "• SPQ Configured: " & IIf(colSPQIdx >= 0, "YES (Auto-verification enabled)", "Standard") & vbCrLf & vbCrLf & _
                   "Next Step: Click [ Save & Create DA ] to save to database.", _
                   vbInformation, "1-Click Parse Succeeded"
        Else
            MsgBox "Header parsed:" & vbCrLf & _
                   "• DA Number: " & wsIn.Range("C4").Value & vbCrLf & _
                   "• Vendor: " & wsIn.Range("E4").Value & vbCrLf & _
                   "• Customer: " & wsIn.Range("G4").Value & vbCrLf & vbCrLf & _
                   "Note: No line items detected matching 'Item No | Make | Desc | Qty | SPQ' pattern." & vbCrLf & _
                   "You can type or paste items directly into rows 17+ below.", _
                   vbInformation, "Header Parsed"
        End If
    End If
    Exit Sub
ErrHandler:
    If Application.Visible Then
        MsgBox "Error parsing email: " & Err.Description, vbCritical, "Parsing Error"
    End If
End Sub

Public Sub SaveDAToDatabase()
    Dim wsIn As Worksheet, wsDB As Worksheet, wsLines As Worksheet
    Set wsIn = ThisWorkbook.Sheets("DA_Intake")
    Set wsDB = ThisWorkbook.Sheets("Daily_Registry")
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    
    Dim daNum As String, vendor As String, cust As String
    daNum = Trim(wsIn.Range("C4").Value)
    vendor = Trim(wsIn.Range("E4").Value)
    cust = Trim(wsIn.Range("G4").Value)
    
    If Len(daNum) = 0 Then
        If Application.Visible Then
            MsgBox "Please parse an email or enter a DA Number in Cell C4 first.", vbExclamation, "No DA Number"
        End If
        Exit Sub
    End If
    
    ' Unprotect lines and database sheets
    wsLines.Unprotect
    wsDB.Unprotect
    
    ' Check if DA exists
    Dim r As Long, lastDB As Long, existsRow As Long
    lastDB = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row
    existsRow = 0
    For r = 5 To lastDB
        If CStr(wsDB.Cells(r, "B").Value) = daNum Then
            existsRow = r
            Exit For
        End If
    Next r
    
    ' Count intake lines
    Dim lastIn As Long, lineCount As Long, totalExp As Long
    lastIn = wsIn.Cells(wsIn.Rows.Count, "C").End(xlUp).Row
    lineCount = 0
    totalExp = 0
    
    ' If DA already exists, ask to overwrite
    If existsRow > 0 Then
        If Application.Visible Then
            Dim respOver As VbMsgBoxResult
            respOver = MsgBox("DA " & daNum & " already exists in the system." & vbCrLf & vbCrLf & _
                              "Do you want to update and overwrite its records?", vbYesNo + vbQuestion, "DA Already Exists")
            If respOver <> vbYes Then
                wsLines.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
                wsDB.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
                Exit Sub
            End If
        End If
        
        ' Remove old lines from DA_Lines
        Dim lastL_old As Long
        lastL_old = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
        For r = lastL_old To 5 Step -1
            If CStr(wsLines.Cells(r, "B").Value) = daNum Then
                wsLines.Rows(r).Delete
            End If
        Next r
    End If
    
    Dim lastL As Long
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    
    For r = 17 To lastIn
        If Len(Trim(wsIn.Cells(r, "C").Value)) > 0 Then
            lineCount = lineCount + 1
            Dim pPart As String, pDesc As String, pQty As Long, pVdr As String, pSPQ As Long
            pPart = UCase(Trim(wsIn.Cells(r, "C").Value))
            pDesc = Trim(wsIn.Cells(r, "D").Value)
            pVdr = Trim(wsIn.Cells(r, "E").Value)
            If Len(pVdr) = 0 Then pVdr = vendor
            pQty = Val(wsIn.Cells(r, "F").Value)
            pSPQ = Val(wsIn.Cells(r, "G").Value)
            totalExp = totalExp + pQty
            
            ' Append to DA_Lines
            lastL = lastL + 1
            wsLines.Cells(lastL, "B").Value = daNum
            wsLines.Cells(lastL, "C").Value = lineCount
            wsLines.Cells(lastL, "D").Value = pPart
            wsLines.Cells(lastL, "E").Value = pDesc
            wsLines.Cells(lastL, "F").Value = pQty
            wsLines.Cells(lastL, "G").Value = 0
            wsLines.Cells(lastL, "H").Value = pQty
            wsLines.Cells(lastL, "I").Value = "PENDING"
            wsLines.Cells(lastL, "J").Value = pSPQ
            wsLines.Cells(lastL, "K").Value = pVdr
        End If
    Next r
    
    If lineCount = 0 Then
        wsLines.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        wsDB.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        If Application.Visible Then
            MsgBox "No line items found to save. Please make sure items are listed in rows 17+ below.", vbExclamation, "No Line Items"
        End If
        Exit Sub
    End If
    
    ' Update or append Daily_Registry
    Dim targetDBRow As Long
    If existsRow > 0 Then
        targetDBRow = existsRow
    Else
        targetDBRow = lastDB + 1
    End If
    
    wsDB.Cells(targetDBRow, "B").Value = daNum
    wsDB.Cells(targetDBRow, "C").Value = Format(Date, "YYYY-MM-DD")
    wsDB.Cells(targetDBRow, "D").Value = cust
    wsDB.Cells(targetDBRow, "E").Value = vendor
    wsDB.Cells(targetDBRow, "F").Value = lineCount
    wsDB.Cells(targetDBRow, "G").Value = totalExp
    wsDB.Cells(targetDBRow, "H").Value = 0
    wsDB.Cells(targetDBRow, "I").Value = totalExp
    wsDB.Cells(targetDBRow, "J").Value = "OPEN"
    
    wsLines.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    wsDB.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    If Application.Visible Then
        MsgBox "DA " & daNum & " successfully saved to database!" & vbCrLf & vbCrLf & _
               "• Added to Daily Registry: Row " & targetDBRow & vbCrLf & _
               "• Line Items Saved: " & lineCount & " lines" & vbCrLf & _
               "• Total Expected Qty: " & Format(totalExp, "#,##0") & " Pcs" & vbCrLf & vbCrLf & _
               "Next Step: Click [ Generate Barcoded DA ] to create the printable barcode slip.", _
               vbInformation, "DA Saved to Database"
    End If
End Sub

Private Function GetCode39Pattern(ByVal c As String) As String
    Select Case UCase(c)
        Case "0": GetCode39Pattern = "000110100"
        Case "1": GetCode39Pattern = "100100001"
        Case "2": GetCode39Pattern = "001100001"
        Case "3": GetCode39Pattern = "101100000"
        Case "4": GetCode39Pattern = "000110001"
        Case "5": GetCode39Pattern = "100110000"
        Case "6": GetCode39Pattern = "001110000"
        Case "7": GetCode39Pattern = "000100101"
        Case "8": GetCode39Pattern = "100100100"
        Case "9": GetCode39Pattern = "001100100"
        Case "A": GetCode39Pattern = "100001001"
        Case "B": GetCode39Pattern = "001001001"
        Case "C": GetCode39Pattern = "101001000"
        Case "D": GetCode39Pattern = "000011001"
        Case "E": GetCode39Pattern = "100011000"
        Case "F": GetCode39Pattern = "001011000"
        Case "G": GetCode39Pattern = "000001101"
        Case "H": GetCode39Pattern = "100001100"
        Case "I": GetCode39Pattern = "001001100"
        Case "J": GetCode39Pattern = "000011100"
        Case "K": GetCode39Pattern = "100000011"
        Case "L": GetCode39Pattern = "001000011"
        Case "M": GetCode39Pattern = "101000010"
        Case "N": GetCode39Pattern = "000010110"
        Case "O": GetCode39Pattern = "100010010"
        Case "P": GetCode39Pattern = "001010010"
        Case "Q": GetCode39Pattern = "000000111"
        Case "R": GetCode39Pattern = "100000110"
        Case "S": GetCode39Pattern = "001000110"
        Case "T": GetCode39Pattern = "000010110"
        Case "U": GetCode39Pattern = "110000001"
        Case "V": GetCode39Pattern = "011000001"
        Case "W": GetCode39Pattern = "111000000"
        Case "X": GetCode39Pattern = "010010001"
        Case "Y": GetCode39Pattern = "110010000"
        Case "Z": GetCode39Pattern = "011010000"
        Case "-": GetCode39Pattern = "010000101"
        Case ".": GetCode39Pattern = "110000100"
        Case " ": GetCode39Pattern = "011000100"
        Case "*": GetCode39Pattern = "010010100"
        Case "$": GetCode39Pattern = "010101000"
        Case "/": GetCode39Pattern = "010100010"
        Case "+": GetCode39Pattern = "010001010"
        Case "%": GetCode39Pattern = "000101010"
        Case Else: GetCode39Pattern = "000110100"
    End Select
End Function

Public Sub DrawDABarcode(ByVal ws As Worksheet, ByVal daText As String)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name Like "DABarcode*" Then shp.Delete
    Next shp
    On Error GoTo 0
    
    daText = Trim(daText)
    If Len(daText) = 0 Then Exit Sub
    
    Dim raw As String
    raw = "*" & UCase(daText) & "*"
    
    Dim narrowW As Double, wideW As Double, gapW As Double, barH As Double
    narrowW = 1.4
    wideW = 3.6
    gapW = 1.4
    barH = 46#
    
    ' Calculate total width of barcode bars
    Dim totalW As Double, i As Long, j As Long, charPat As String, bitVal As String
    totalW = 0#
    For i = 1 To Len(raw)
        charPat = GetCode39Pattern(Mid(raw, i, 1))
        For j = 1 To 9
            If Mid(charPat, j, 1) = "1" Then
                totalW = totalW + wideW
            Else
                totalW = totalW + narrowW
            End If
        Next j
        totalW = totalW + gapW
    Next i
    totalW = totalW - gapW
    
    ' Card Box dimensions
    Dim cardPadX As Double, cardPadY As Double
    cardPadX = 26#
    cardPadY = 8#
    Dim cardW As Double, cardH As Double
    cardW = totalW + (cardPadX * 2)
    cardH = barH + 30# + (cardPadY * 2)
    
    ' Center card horizontally across columns B to G
    Dim gridLeft As Double, gridW As Double
    gridLeft = ws.Range("B4").Left
    gridW = ws.Range("G4").Left + ws.Range("G4").Width - gridLeft
    
    Dim cardLeft As Double, cardTop As Double
    cardLeft = gridLeft + (gridW - cardW) / 2#
    cardTop = ws.Range("B4").Top + 4#
    
    ' Draw Card Background
    Dim bgCard As Shape
    Set bgCard = ws.Shapes.AddShape(5, cardLeft, cardTop, cardW, cardH) ' 5 = msoShapeRoundedRectangle
    bgCard.Fill.Solid
    bgCard.Fill.ForeColor.RGB = RGB(255, 255, 255)
    bgCard.Line.Weight = 1.2
    bgCard.Line.ForeColor.RGB = RGB(203, 213, 225) ' Slate 300
    bgCard.Name = "DABarcode_BG"
    
    Dim shapeNames() As String
    Dim shapeCount As Long
    shapeCount = 1
    ReDim shapeNames(1 To 250)
    shapeNames(1) = bgCard.Name
    
    Dim curX As Double, barTop As Double
    curX = cardLeft + cardPadX
    barTop = cardTop + cardPadY + 4#
    
    Dim bIdx As Long
    bIdx = 0
    For i = 1 To Len(raw)
        charPat = GetCode39Pattern(Mid(raw, i, 1))
        For j = 1 To 9
            bitVal = Mid(charPat, j, 1)
            Dim elemW As Double
            If bitVal = "1" Then elemW = wideW Else elemW = narrowW
            
            ' Odd index (1, 3, 5, 7, 9) is a BAR
            If (j Mod 2) = 1 Then
                bIdx = bIdx + 1
                Dim barShp As Shape
                Set barShp = ws.Shapes.AddShape(1, curX, barTop, elemW, barH) ' 1 = msoShapeRectangle
                barShp.Fill.Solid
                barShp.Fill.ForeColor.RGB = RGB(0, 0, 0)
                barShp.Line.Visible = 0 ' msoFalse
                shapeCount = shapeCount + 1
                If shapeCount > UBound(shapeNames) Then ReDim Preserve shapeNames(1 To shapeCount + 50)
                barShp.Name = "DABarcode_B" & bIdx
                shapeNames(shapeCount) = barShp.Name
            End If
            curX = curX + elemW
        Next j
        curX = curX + gapW
    Next i
    
    ' Human readable text label below bars
    Dim lblShp As Shape
    Set lblShp = ws.Shapes.AddTextbox(1, cardLeft, barTop + barH + 4#, cardW, 20#)
    lblShp.Fill.Visible = 0
    lblShp.Line.Visible = 0
    lblShp.TextFrame2.MarginLeft = 0
    lblShp.TextFrame2.MarginRight = 0
    lblShp.TextFrame2.MarginTop = 0
    lblShp.TextFrame2.MarginBottom = 0
    lblShp.TextFrame2.WordWrap = 0
    lblShp.TextFrame2.TextRange.Text = "*DA" & daText & "*   [ SCAN DA BARCODE ]"
    lblShp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    lblShp.TextFrame2.TextRange.Font.Size = 9#
    lblShp.TextFrame2.TextRange.Font.Bold = -1 ' msoTrue
    lblShp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(15, 23, 42)
    lblShp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter
    shapeCount = shapeCount + 1
    If shapeCount > UBound(shapeNames) Then ReDim Preserve shapeNames(1 To shapeCount + 10)
    lblShp.Name = "DABarcode_LBL"
    shapeNames(shapeCount) = lblShp.Name
    
    ' Group all barcode elements
    ReDim Preserve shapeNames(1 To shapeCount)
    Dim grp As Shape
    Set grp = ws.Shapes.Range(shapeNames).Group
    grp.Name = "DABarcodeGroup"
    grp.Placement = 3
    grp.Locked = True
End Sub

Public Sub PrintDASlip()
    ActiveSheet.PrintPreview
End Sub

Public Sub OpenDAInConsole()
    Dim daNum As String
    daNum = Trim(ThisWorkbook.Sheets("Printable_DA").Range("C8").Value)
    If Len(daNum) > 0 Then
        modQCLogic.LoadDAConsole daNum
        ThisWorkbook.Sheets("QC_Console").Activate
    End If
End Sub

Public Sub ExportBarcodeImage()
    Dim wsPrint As Worksheet
    Set wsPrint = ThisWorkbook.Sheets("Printable_DA")
    
    Dim daNum As String
    daNum = Trim(wsPrint.Range("C8").Value)
    If Len(daNum) = 0 Then daNum = "Current"
    
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsPrint.Shapes("DABarcodeGroup")
    On Error GoTo 0
    
    If shp Is Nothing Then
        MsgBox "Barcode shape not found on Printable_DA sheet.", vbExclamation, "Export Barcode"
        Exit Sub
    End If
    
    Dim outPath As String
    If Len(ThisWorkbook.Path) > 0 Then
        outPath = ThisWorkbook.Path & "\DA_" & daNum & "_Barcode.png"
    Else
        outPath = Environ("USERPROFILE") & "\Desktop\DA_" & daNum & "_Barcode.png"
    End If
    
    ' Copy shape as picture and export via temporary chart
    shp.CopyPicture Appearance:=1, Format:=2
    Dim cht As ChartObject
    Set cht = wsPrint.ChartObjects.Add(0, 0, shp.Width, shp.Height)
    cht.Chart.Paste
    cht.Chart.Export Filename:=outPath, FilterName:="PNG"
    cht.Delete
    
    MsgBox "Barcode image successfully generated and saved to:" & vbCrLf & vbCrLf & outPath & vbCrLf & vbCrLf & "Ready for use in label printing, packing slips, emails, or carton tags.", vbInformation, "Barcode Image Exported"
End Sub

Public Sub GeneratePrintableDA()
    Dim wsIn As Worksheet, wsPrint As Worksheet
    Set wsIn = ThisWorkbook.Sheets("DA_Intake")
    Set wsPrint = ThisWorkbook.Sheets("Printable_DA")
    
    Dim daNum As String, vendor As String, cust As String
    daNum = Trim(wsIn.Range("C4").Value)
    vendor = Trim(wsIn.Range("E4").Value)
    cust = Trim(wsIn.Range("G4").Value)
    
    If Len(daNum) = 0 Then daNum = "27006862"
    
    ' Clear old text from barcode rows
    wsPrint.Range("B4:G7").ClearContents
    
    ' Render authentic vector barcode card
    DrawDABarcode wsPrint, daNum
    
    wsPrint.Range("C8").Value = daNum
    wsPrint.Range("E8").Value = Format(Date, "YYYY-MM-DD")
    wsPrint.Range("C9").Value = cust
    wsPrint.Range("E9").Value = vendor
    
    ' Auto-export image file to workbook directory
    On Error Resume Next
    If Len(ThisWorkbook.Path) > 0 Then
        Dim autoImgPath As String
        autoImgPath = ThisWorkbook.Path & "\DA_" & daNum & "_Barcode.png"
        Dim sGrp As Shape
        Set sGrp = wsPrint.Shapes("DABarcodeGroup")
        If Not sGrp Is Nothing Then
            sGrp.CopyPicture Appearance:=1, Format:=2
            Dim chAuto As ChartObject
            Set chAuto = wsPrint.ChartObjects.Add(0, 0, sGrp.Width, sGrp.Height)
            chAuto.Chart.Paste
            chAuto.Chart.Export Filename:=autoImgPath, FilterName:="PNG"
            chAuto.Delete
        End If
    End If
    On Error GoTo 0
    
    ' Copy intake lines
    wsPrint.Range("B13:H60").ClearContents
    Dim r As Long, outR As Long, lastIn As Long
    lastIn = wsIn.Cells(wsIn.Rows.Count, "C").End(xlUp).Row
    outR = 13
    For r = 17 To lastIn
        If Len(Trim(wsIn.Cells(r, "C").Value)) > 0 Then
            wsPrint.Cells(outR, "B").Value = wsIn.Cells(r, "B").Value
            wsPrint.Cells(outR, "C").Value = wsIn.Cells(r, "C").Value
            wsPrint.Cells(outR, "D").Value = wsIn.Cells(r, "D").Value
            wsPrint.Cells(outR, "E").Value = wsIn.Cells(r, "F").Value
            wsPrint.Cells(outR, "F").Value = wsIn.Cells(r, "G").Value
            wsPrint.Cells(outR, "G").Value = "[   ]"
            wsPrint.Cells(outR, "H").Value = ""
            outR = outR + 1
        End If
    Next r
    
    ' Apply WrapText and responsive row heights to Printable_DA table
    If outR > 13 Then
        Dim maxPrR As Long: maxPrR = outR - 1
        With wsPrint.Range("B13:H" & maxPrR)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        Dim rPrAdj As Long
        For rPrAdj = 13 To maxPrR
            If wsPrint.Rows(rPrAdj).RowHeight < 24 Then wsPrint.Rows(rPrAdj).RowHeight = 24
        Next rPrAdj
    End If
    
    wsPrint.Activate
End Sub

Public Sub ClearIntakeStaging()
    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Sheets("DA_Intake")
    wsIn.Unprotect
    wsIn.Range("B7").Value = "Click [ OPEN EMAIL INTAKE FORM ] below to paste and extract Outlook DA emails."
    wsIn.Range("C4").Value = ""
    wsIn.Range("E4").Value = ""
    wsIn.Range("G4").Value = ""
    
    Dim lastR As Long
    lastR = wsIn.Cells(wsIn.Rows.Count, "B").End(xlUp).Row
    If lastR >= 17 Then
        wsIn.Range("B17:H" & lastR).ClearContents
        wsIn.Range("B17:H" & lastR).Interior.ColorIndex = xlNone
    End If
    wsIn.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    If Application.Visible Then
        MsgBox "Intake staging table cleared. Ready for new DA email.", vbInformation, "Staging Cleared"
    End If
End Sub

Public Sub ShowDAIntakeForm()
    On Error Resume Next
    frmDAEmailIntake.Show vbModal
    On Error GoTo 0
End Sub

Public Sub ShowReprintForm()
    On Error Resume Next
    frmReprintDA.Show vbModal
    On Error GoTo 0
End Sub

Public Sub LoadHistoricalDAToPrint(ByVal targetDA As String)
    Dim wsPrint As Worksheet, wsDB As Worksheet, wsLines As Worksheet
    Set wsPrint = ThisWorkbook.Sheets("Printable_DA")
    Set wsDB = ThisWorkbook.Sheets("Daily_Registry")
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    
    If Len(targetDA) = 0 Then Exit Sub
    
    Dim r As Long, lastDB As Long, found As Boolean
    Dim cust As String, vendor As String, daDate As String, daStatus As String
    lastDB = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row
    found = False
    
    For r = 5 To lastDB
        If CStr(wsDB.Cells(r, "B").Value) = targetDA Then
            daDate = Format(wsDB.Cells(r, "C").Value, "YYYY-MM-DD")
            cust = CStr(wsDB.Cells(r, "D").Value)
            vendor = CStr(wsDB.Cells(r, "E").Value)
            daStatus = CStr(wsDB.Cells(r, "J").Value)
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        If Application.Visible Then
            MsgBox "DA " & targetDA & " was not found in Daily_Registry.", vbExclamation, "DA Not Found"
        End If
        Exit Sub
    End If
    
    wsPrint.Unprotect
    wsPrint.Range("C8").Value = targetDA
    wsPrint.Range("E8").Value = daDate
    wsPrint.Range("G8").Value = daStatus
    wsPrint.Range("C9").Value = cust
    wsPrint.Range("E9").Value = vendor
    
    DrawDABarcode wsPrint, targetDA
    
    Dim lastL As Long, outRow As Long
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    outRow = 13
    
    wsPrint.Range("B13:H50").ClearContents
    wsPrint.Range("B13:H50").Interior.ColorIndex = xlNone
    
    Dim itemIdx As Long: itemIdx = 0
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = targetDA Then
            itemIdx = itemIdx + 1
            wsPrint.Cells(outRow, "B").Value = itemIdx
            wsPrint.Cells(outRow, "C").Value = wsLines.Cells(r, "D").Value
            wsPrint.Cells(outRow, "D").Value = wsLines.Cells(r, "E").Value
            wsPrint.Cells(outRow, "E").Value = wsLines.Cells(r, "F").Value
            wsPrint.Cells(outRow, "F").Value = wsLines.Cells(r, "J").Value
            wsPrint.Cells(outRow, "G").Value = IIf(wsLines.Cells(r, "I").Value = "MATCHED", "[ VERIFIED ]", "[   ]")
            wsPrint.Cells(outRow, "H").Value = ""
            outRow = outRow + 1
        End If
    Next r
    
    wsPrint.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    wsPrint.Activate
End Sub


