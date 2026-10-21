VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQCScanner 
   Caption         =   "Warehouse QC Scanning Terminal"
   ClientHeight    =   8220.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10560
   OleObjectBlob   =   "frmQCScanner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQCScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    lblBanner.Caption = "READY — SCAN DA BARCODE"
    lblBanner.BackColor = RGB(2, 132, 199)
    txtScanDA.Text = "27006862"
    LoadActiveDA
End Sub

Private Sub txtScanDA_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        LoadActiveDA
        txtScanComponent.SetFocus
    End If
End Sub

Private Sub btnLoadDA_Click()
    LoadActiveDA
    txtScanComponent.SetFocus
End Sub

Private Sub LoadActiveDA()
    Dim daNum As String
    daNum = Trim(txtScanDA.Text)
    If Len(daNum) = 0 Then Exit Sub
    
    modQCLogic.LoadDAConsole daNum
    
    ' Refresh UserForm display
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    lblDAInfo.Caption = "DA: " & wsQC.Range("C10").Value & "  |  Vendor: " & wsQC.Range("F10").Value & "  |  Customer: " & wsQC.Range("H10").Value
    
    ' Populate ListBox
    lstItems.Clear
    Dim r As Long
    For r = 16 To 45
        If Len(Trim(wsQC.Cells(r, "C").Value)) > 0 And Trim(wsQC.Cells(r, "C").Value) <> "TOTALS:" Then
            lstItems.AddItem wsQC.Cells(r, "C").Value & "  |  Exp: " & wsQC.Cells(r, "F").Value & "  |  Checked: " & wsQC.Cells(r, "G").Value & "  |  Rem: " & wsQC.Cells(r, "H").Value & "  [" & wsQC.Cells(r, "J").Value & "]"
        End If
    Next r
    
    lblBanner.Caption = "DA " & daNum & " LOADED — SCAN COMPONENT"
    lblBanner.BackColor = RGB(2, 132, 199)
End Sub

Private Sub txtScanComponent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Dim vName As String, rawScan As String, cleanP As String, daNum As String, itemSPQ As Long
        daNum = Trim(txtScanDA.Text)
        vName = Trim(ThisWorkbook.Sheets("QC_Console").Range("F10").Value)
        rawScan = Trim(txtScanComponent.Text)
        cleanP = modQCLogic.CleanPartNumber(rawScan, vName)
        lblCleanPart.Caption = cleanP
        
        ' Check if SPQ is mentioned in DA
        itemSPQ = modQCLogic.GetItemSPQ(daNum, cleanP)
        If itemSPQ > 0 Then
            ' ZERO-TOUCH: auto-populate quantity and auto-submit without user input
            txtQuantity.Text = CStr(itemSPQ)
            btnProcessScan_Click
        Else
            txtQuantity.SetFocus
        End If
    End If
End Sub

Private Sub txtQuantity_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        btnProcessScan_Click
    End If
End Sub

Private Sub btnProcessScan_Click()
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    
    wsQC.Range("C12").Value = txtScanComponent.Text
    wsQC.Range("G12").Value = txtQuantity.Text
    
    modQCLogic.ProcessScanConsole
    
    ' Update UserForm
    lblBanner.Caption = wsQC.Range("B3").Value
    lblBanner.BackColor = wsQC.Range("B3").Interior.Color
    
    ' Refresh ListBox
    LoadActiveDA
    
    txtScanComponent.Text = ""
    txtQuantity.Text = ""
    lblCleanPart.Caption = "-"
    txtScanComponent.SetFocus
End Sub

Private Sub btnDeleteScan_Click()
    Dim rawC As String
    rawC = Trim(txtScanComponent.Text)
    If Len(rawC) > 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("CONFIRM DELETE SCANNED BARCODE:" & vbCrLf & vbCrLf & _
                      "• Scanned Barcode: " & rawC & vbCrLf & _
                      "• Cleaned Part: " & lblCleanPart.Caption & vbCrLf & _
                      IIf(Len(txtQuantity.Text) > 0, "• Quantity: " & txtQuantity.Text & " Pcs" & vbCrLf, "") & vbCrLf & _
                      "Do you want to delete and clear this barcode from the scanner?", _
                      vbYesNo + vbExclamation, "Delete Scanned Barcode")
        If resp = vbYes Then
            txtScanComponent.Text = ""
            txtQuantity.Text = ""
            lblCleanPart.Caption = "-"
            modQCLogic.ExecuteDeleteCurrentScan
            lblBanner.Caption = "SCANNED BARCODE DELETED — READY FOR NEXT SCAN"
            lblBanner.BackColor = RGB(2, 132, 199)
            txtScanComponent.SetFocus
        End If
    Else
        modQCLogic.DeleteOrClearScan
        LoadActiveDA
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


