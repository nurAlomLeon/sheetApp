VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuantityInput 
   Caption         =   "Step 3: Component Quantity Verification"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8952.001
   OleObjectBlob   =   "frmQuantityInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuantityInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_targetPart As String
Private m_expQty As Long
Private m_scnQty As Long
Private m_remQty As Long
Private m_spq As Long

Public Sub SetupPart(ByVal partNo As String)
    m_targetPart = partNo
End Sub

Private Sub UserForm_Initialize()
    lblBanner.Caption = "STEP 3: QUANTITY VERIFICATION"
    lblBanner.BackColor = RGB(217, 119, 6)
    
    Dim wsQC As Worksheet, wsLines As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    Set wsLines = ThisWorkbook.Sheets("DA_Lines")
    
    Dim daNum As String, curPart As String
    daNum = Trim(wsQC.Range("C10").Value)
    curPart = Trim(wsQC.Range("C12").Value)
    If Len(m_targetPart) = 0 Then m_targetPart = curPart
    
    lblPartVal.Caption = m_targetPart
    
    Dim r As Long, lastL As Long, found As Boolean
    lastL = wsLines.Cells(wsLines.Rows.Count, "B").End(xlUp).Row
    found = False
    
    For r = 5 To lastL
        If CStr(wsLines.Cells(r, "B").Value) = daNum And UCase(Trim(wsLines.Cells(r, "D").Value)) = UCase(m_targetPart) Then
            m_expQty = Val(wsLines.Cells(r, "F").Value)
            m_scnQty = Val(wsLines.Cells(r, "G").Value)
            m_remQty = Val(wsLines.Cells(r, "H").Value)
            m_spq = Val(wsLines.Cells(r, "J").Value)
            lblDescVal.Caption = Trim(wsLines.Cells(r, "E").Value)
            lblExpVal.Caption = Format(m_expQty, "#,##0") & " Pcs"
            lblScnVal.Caption = Format(m_scnQty, "#,##0") & " Pcs"
            lblRemVal.Caption = Format(m_remQty, "#,##0") & " Pcs Remaining"
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        lblDescVal.Caption = "-"
        lblExpVal.Caption = "0 Pcs"
        lblScnVal.Caption = "0 Pcs"
        lblRemVal.Caption = "0 Pcs"
    End If
    
    If m_spq > 0 Then
        txtQuantity.Text = CStr(m_spq)
    Else
        txtQuantity.Text = ""
    End If
    
    txtQuantity.SetFocus
End Sub

Private Sub btnSubmitQty_Click()
    Dim rawQ As String
    rawQ = Trim(txtQuantity.Text)
    If Len(rawQ) = 0 Then
        MsgBox "Please scan or type the carton quantity.", vbExclamation, "Quantity Missing"
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    Dim cleanQ As Long
    cleanQ = modQCLogic.CleanQuantity(rawQ)
    
    If cleanQ <= 0 Then
        MsgBox "Please enter a valid positive quantity.", vbExclamation, "Invalid Quantity"
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    If cleanQ > m_remQty Then
        MsgBox "QUANTITY EXCEEDS OUTSTANDING DA QUANTITY!" & vbCrLf & vbCrLf & _
               "• Outstanding Remaining Balance: " & Format(m_remQty, "#,##0") & " Pcs" & vbCrLf & _
               "• Entered Quantity: " & Format(cleanQ, "#,##0") & " Pcs" & vbCrLf & _
               "• Excess: +" & Format(cleanQ - m_remQty, "#,##0") & " Pcs" & vbCrLf & vbCrLf & _
               "Transaction rejected to prevent over-packing.", vbCritical, "Over-Quantity Rejected"
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    wsQC.Unprotect
    wsQC.Range("G12").Value = cleanQ
    wsQC.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    Unload Me
    modQCLogic.ProcessScanConsole
End Sub

Private Sub btnDeleteScan_Click()
    modQCLogic.ExecuteDeleteCurrentScan
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


