VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSPQInput 
   Caption         =   "Standard Pack Quantity (SPQ) Verification"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8952.001
   OleObjectBlob   =   "frmSPQInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSPQInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    
    Dim daNum As String, vendor As String, cust As String
    daNum = Trim(wsQC.Range("C10").Value)
    vendor = Trim(wsQC.Range("F10").Value)
    cust = Trim(wsQC.Range("H10").Value)
    
    lblDAInfo.Caption = "Active DA: " & daNum & "  |  Vendor: " & vendor & "  |  Customer: " & cust
    
    Dim compPart As String, cleanPart As String
    compPart = Trim(wsQC.Range("C12").Value)
    If Len(compPart) > 0 Then
        cleanPart = modQCLogic.CleanPartNumber(compPart, vendor)
    Else
        cleanPart = Trim(wsQC.Range("C16").Value)
    End If
    If Len(cleanPart) = 0 Then cleanPart = "BLM18KG121TH1D"
    
    lblPartInfo.Caption = "Component Part: " & cleanPart
    
    Dim wsLines As Worksheet
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    Dim r As Long, lastL As Long, remQty As Long, expQty As Long
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum And UCase(Trim(wsLines.Cells(r, "D").Value)) = UCase(cleanPart) Then
            expQty = Val(wsLines.Cells(r, "F").Value)
            remQty = Val(wsLines.Cells(r, "H").Value)
            Exit For
        End If
    Next r
    
    lblBalInfo.Caption = "Target Order Qty: " & Format(expQty, "#,##0") & " Pcs  |  Remaining: " & Format(remQty, "#,##0") & " Pcs"
    
    Dim knownSPQ As Long
    knownSPQ = modQCLogic.GetItemSPQ(daNum, cleanPart)
    If knownSPQ > 0 Then
        txtSPQValue.Text = CStr(knownSPQ)
        lblPrompt.Caption = "DA Configured SPQ (" & Format(knownSPQ, "#,##0") & " Pcs) detected. Confirm or override:"
    Else
        txtSPQValue.Text = ""
        lblPrompt.Caption = "Enter Standard Pack Quantity (SPQ) to register:"
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSubmitSPQ_Click()
    Dim rawInput As String
    rawInput = Trim(txtSPQValue.Text)
    rawInput = Replace(rawInput, ",", "")
    rawInput = Replace(rawInput, " ", "")
    
    If Not IsNumeric(rawInput) Or Val(rawInput) <= 0 Then
        MsgBox "Please enter a valid numeric Standard Pack Quantity (greater than 0).", vbCritical, "Invalid SPQ"
        txtSPQValue.SetFocus
        Exit Sub
    End If
    
    Dim spqVal As Long
    spqVal = CLng(rawInput)
    
    Dim partStr As String
    partStr = Mid(lblPartInfo.Caption, 17)
    
    ' ========================================================
    ' STEP 1 OF 2 VERIFICATION: First Confirmation Dialog
    ' ========================================================
    Dim resp1 As VbMsgBoxResult
    resp1 = MsgBox("STEP 1 OF 2 VERIFICATION:" & vbCrLf & vbCrLf & _
                   "Target Component Part: " & partStr & vbCrLf & _
                   "Entered SPQ Quantity: " & Format(spqVal, "#,##0") & " Pcs" & vbCrLf & vbCrLf & _
                   "Is this Standard Pack Quantity correct?", _
                   vbYesNo + vbQuestion, "Step 1 of 2: Verify SPQ")
                   
    If resp1 <> vbYes Then
        MsgBox "SPQ entry cancelled by operator. Nothing was registered.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' ========================================================
    ' STEP 2 OF 2 VERIFICATION: Final Safety Double-Check
    ' ========================================================
    Dim resp2 As VbMsgBoxResult
    resp2 = MsgBox("FINAL STEP 2 CONFIRMATION:" & vbCrLf & vbCrLf & _
                   "PLEASE DOUBLE-CHECK THE PHYSICAL CARTON / REEL LABEL!" & vbCrLf & vbCrLf & _
                   "Are you ABSOLUTELY SURE you want to register " & Format(spqVal, "#,##0") & " Pcs into the system?" & vbCrLf & _
                   "This action will update the warehouse records and audit log.", _
                   vbOKCancel + vbExclamation, "Step 2 of 2: Final Safety Check")
                   
    If resp2 <> vbOK Then
        MsgBox "Final safety confirmation rejected. Nothing was registered.", vbInformation, "Aborted"
        Exit Sub
    End If
    
    ' DOUBLE CONFIRMATION COMPLETE: Apply to QC Console
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    modQCLogic.UnprotectConsole
    If Len(Trim(wsQC.Range("C12").Value)) = 0 Then
        wsQC.Range("C12").Value = partStr
    End If
    wsQC.Range("G12").Value = spqVal
    modQCLogic.ProtectConsole
    
    ' Process through deterministic verification engine
    modQCLogic.ProcessScanConsole
    
    Unload Me
End Sub


