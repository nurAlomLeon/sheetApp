VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Dim todayStr As String
    todayStr = Format(Date, "DD-MMM-YYYY (dddd)")
    
    modQCLogic.UnprotectConsole
    
    ' 1. Recognize day on QC_Console
    Dim wsQC As Worksheet
    Set wsQC = ThisWorkbook.Sheets("QC_Console")
    wsQC.Range("B2").Value = "WAREHOUSE QC AUTOMATION SYSTEM — TODAY: " & UCase(todayStr)
    wsQC.Range("B3:J4").Value = "DAILY QC SESSION: " & UCase(todayStr) & " — READY TO SCAN"
    wsQC.Range("B3:J4").Interior.Color = RGB(2, 132, 199)
    
    ' 2. Recognize day on Daily_Registry
    Dim wsReg As Worksheet
    Set wsReg = ThisWorkbook.Sheets("Daily_Registry")
    wsReg.Range("B2").Value = "CENTRAL DAILY REGISTRY — RECOGNIZED DATE: " & UCase(todayStr)
    
    modQCLogic.ProtectAllSheets
    
    ' 3. Activate QC_Console and select Box 1
    wsQC.Activate
    wsQC.Range("C10").Select
    
    Application.StatusBar = "Session Started: " & todayStr & " | Central Daily Registry Active."
End Sub

