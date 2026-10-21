VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReprintDA 
   Caption         =   "Historical Barcode Retrieval & Reprint"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9960.001
   OleObjectBlob   =   "frmReprintDA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReprintDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    lblBanner.Caption = "HISTORICAL BARCODE RETRIEVAL & REPRINT"
    lblBanner.BackColor = RGB(79, 70, 229)
    
    lstDAs.Clear
    
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("Daily_Registry")
    
    Dim r As Long, lastDB As Long
    lastDB = wsDB.Cells(wsDB.Rows.Count, "B").End(xlUp).Row
    
    For r = 5 To lastDB
        Dim daVal As String, custVal As String, statusVal As String, dateVal As String, qtyVal As String
        daVal = Trim(CStr(wsDB.Cells(r, "B").Value))
        If Len(daVal) > 0 Then
            dateVal = Format(wsDB.Cells(r, "C").Value, "YYYY-MM-DD")
            custVal = Trim(CStr(wsDB.Cells(r, "D").Value))
            qtyVal = Format(Val(wsDB.Cells(r, "G").Value), "#,##0") & " Pcs"
            statusVal = Trim(CStr(wsDB.Cells(r, "J").Value))
            lstDAs.AddItem daVal & "  |  " & custVal & "  |  " & dateVal & "  |  " & qtyVal & "  |  " & statusVal
        End If
    Next r
    
    If lstDAs.ListCount > 0 Then
        lstDAs.ListIndex = 0
        btnLoad.Enabled = True
    Else
        btnLoad.Enabled = False
    End If
End Sub

Private Sub btnLoad_Click()
    If lstDAs.ListIndex < 0 Then
        MsgBox "Please select a DA from the list.", vbExclamation, "No DA Selected"
        Exit Sub
    End If
    
    Dim selText As String, daNum As String, spPos As Long
    selText = lstDAs.List(lstDAs.ListIndex)
    spPos = InStr(selText, " ")
    If spPos > 0 Then daNum = Trim(Left(selText, spPos - 1)) Else daNum = Trim(selText)
    
    modIntakeLogic.LoadHistoricalDAToPrint daNum
    Unload Me
    
    MsgBox "Historical DA " & daNum & " successfully retrieved and loaded onto Printable_DA sheet." & vbCrLf & vbCrLf & _
           "The vector Code 39 barcode and all line items are ready for printing or exporting.", vbInformation, "DA Loaded"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


