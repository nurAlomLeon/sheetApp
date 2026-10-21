VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDAEmailIntake 
   Caption         =   "Register New Dispatch Advice (DA) — Email Ingestion Terminal"
   ClientHeight    =   8220.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10752
   OleObjectBlob   =   "frmDAEmailIntake.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDAEmailIntake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    lblBanner.Caption = "OUTLOOK DA EMAIL INGESTION — ZERO WORKSHEET RISK"
    lblBanner.BackColor = RGB(2, 132, 199)
    txtEmailText.Text = ""
    lblExtractedDA.Caption = "DA Number: -"
    lblExtractedCust.Caption = "Customer: -"
    lblExtractedItems.Caption = "Items Extracted: 0"
    lblExtractedQty.Caption = "Total Qty: 0 Pcs"
    btnSave.Enabled = False
End Sub

Private Sub btnParse_Click()
    Dim rawT As String
    rawT = Trim(txtEmailText.Text)
    If Len(rawT) = 0 Then
        MsgBox "Please paste the raw Outlook DA email text or table into the box first.", vbExclamation, "No Text Found"
        txtEmailText.SetFocus
        Exit Sub
    End If
    
    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Sheets("DA_Intake")
    wsIn.Unprotect
    wsIn.Range("B7").Value = rawT
    wsIn.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    modIntakeLogic.ParseDAEmailText
    
    Dim daNum As String, cust As String, lastR As Long, totQ As Long, itemCnt As Long, r As Long
    daNum = Trim(wsIn.Range("C4").Value)
    cust = Trim(wsIn.Range("G4").Value)
    
    lastR = wsIn.Cells(wsIn.Rows.Count, "B").End(xlUp).Row
    itemCnt = 0
    totQ = 0
    If lastR >= 17 Then
        For r = 17 To lastR
            If Len(Trim(wsIn.Cells(r, "C").Value)) > 0 Then
                itemCnt = itemCnt + 1
                totQ = totQ + Val(wsIn.Cells(r, "F").Value)
            End If
        Next r
    End If
    
    If Len(daNum) > 0 And itemCnt > 0 Then
        lblExtractedDA.Caption = "DA Number: " & daNum
        lblExtractedCust.Caption = "Customer: " & cust
        lblExtractedItems.Caption = "Items Extracted: " & itemCnt & " Lines"
        lblExtractedQty.Caption = "Total Qty: " & Format(totQ, "#,##0") & " Pcs"
        btnSave.Enabled = True
        btnSave.SetFocus
        MsgBox "Email successfully parsed!" & vbCrLf & vbCrLf & _
               "• DA Number: " & daNum & vbCrLf & _
               "• Customer: " & cust & vbCrLf & _
               "• Total Items: " & itemCnt & " Line Items" & vbCrLf & _
               "• Total Dispatch Qty: " & Format(totQ, "#,##0") & " Pcs" & vbCrLf & vbCrLf & _
               "Click [ SAVE & REGISTER DA ] to register into master registry.", vbInformation, "Parsing Successful"
    Else
        MsgBox "Could not extract valid DA line items. Please verify the email table format.", vbExclamation, "Parsing Failed"
    End If
End Sub

Private Sub btnSave_Click()
    modIntakeLogic.SaveDAToDatabase
    Unload Me
End Sub

Private Sub btnClear_Click()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Are you sure you want to clear the current email input and staged lines?" & vbCrLf & vbCrLf & _
                  "(This will NOT delete any previously saved historical DAs).", vbYesNo + vbQuestion, "Clear Email Intake")
    If resp = vbYes Then
        txtEmailText.Text = ""
        lblExtractedDA.Caption = "DA Number: -"
        lblExtractedCust.Caption = "Customer: -"
        lblExtractedItems.Caption = "Items Extracted: 0"
        lblExtractedQty.Caption = "Total Qty: 0 Pcs"
        btnSave.Enabled = False
        modIntakeLogic.ClearIntakeStaging
        txtEmailText.SetFocus
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


