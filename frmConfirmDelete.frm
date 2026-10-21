VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfirmDelete 
   Caption         =   "Confirm Delete Scanned Barcode"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8952.001
   OleObjectBlob   =   "frmConfirmDelete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfirmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    
    Dim rawScan As String, cleanPart As String, rawQty As String, daNum As String
    daNum = Trim(wsQC.Range("C10").Value)
    rawScan = Trim(wsQC.Range("C12").Value)
    cleanPart = modQCLogic.CleanPartNumber(rawScan, Trim(wsQC.Range("F10").Value))
    rawQty = Trim(wsQC.Range("G12").Value)
    
    lblActiveDA.Caption = "Active DA: " & daNum & "  |  Vendor: " & wsQC.Range("F10").Value
    
    If Len(rawScan) > 0 Then
        lblBarcodeVal.Caption = rawScan
        lblPartVal.Caption = IIf(Len(cleanPart) > 0, cleanPart, rawScan)
        lblQtyVal.Caption = IIf(Len(rawQty) > 0, Format(Val(rawQty), "#,##0") & " Pcs", "Not entered / Auto SPQ")
        
        lblStatusMsg.Caption = "Wrong barcode scanned? Click [ DELETE & CLEAR SCAN ] to remove it without touching the cells."
        btnDeleteCurrent.Enabled = True
        btnDeleteCurrent.SetFocus
    Else
        lblBarcodeVal.Caption = "(Scan box [2] is currently empty)"
        lblPartVal.Caption = "-"
        lblQtyVal.Caption = "-"
        
        Dim lastPart As String, lastQty As Long
        modQCLogic.GetLastVerifiedScan daNum, lastPart, lastQty
        
        If Len(lastPart) > 0 And lastQty > 0 Then
            lblStatusMsg.Caption = "Recent verified scan found: " & lastPart & " (" & Format(lastQty, "#,##0") & " Pcs). Click [ UNDO LAST SCAN ] to revert it."
            btnUndoLast.Caption = "UNDO LAST SCAN (" & lastPart & ")"
            btnUndoLast.Enabled = True
            btnUndoLast.SetFocus
        Else
            lblStatusMsg.Caption = "No active barcode in Box [2] and no recent scans found to undo."
            btnUndoLast.Enabled = False
            btnCancel.SetFocus
        End If
        btnDeleteCurrent.Enabled = False
    End If
End Sub

Private Sub btnDeleteCurrent_Click()
    modQCLogic.ExecuteDeleteCurrentScan
    Unload Me
End Sub

Private Sub btnUndoLast_Click()
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    Dim daNum As String
    daNum = Trim(wsQC.Range("C10").Value)
    
    Dim lastPart As String, lastQty As Long
    modQCLogic.GetLastVerifiedScan daNum, lastPart, lastQty
    
    If Len(lastPart) > 0 And lastQty > 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("CONFIRM SCAN ROLLBACK:" & vbCrLf & vbCrLf & _
                      "• Component Part: " & lastPart & vbCrLf & _
                      "• Quantity to Revert: " & Format(lastQty, "#,##0") & " Pcs" & vbCrLf & vbCrLf & _
                      "Are you sure you want to delete and UNDO this verified scan?", _
                      vbYesNo + vbExclamation, "Confirm Rollback")
        If resp = vbYes Then
            modQCLogic.ExecuteUndoLastScan daNum, lastPart, lastQty
        End If
    End If
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


