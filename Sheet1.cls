VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    Application.EnableEvents = False
    
    ' If user changed DA barcode cell (C10), auto-load DA
    If Not Intersect(Target, Me.Range("C10")) Is Nothing Then
        If Len(Trim(Me.Range("C10").Value)) > 0 Then
            modQCLogic.LoadDAConsole Trim(Me.Range("C10").Value)
        End If
    End If
    
    ' If user scanned component (C12)
    If Not Intersect(Target, Me.Range("C12")) Is Nothing Then
        Dim rawScan As String
        rawScan = Trim(Me.Range("C12").Value)
        If Len(rawScan) > 0 Then
            Dim daNum As String, vName As String, cleanPart As String, itemSPQ As Long
            daNum = Trim(Me.Range("C10").Value)
            vName = Trim(Me.Range("F10").Value)
            
            cleanPart = modQCLogic.CleanPartNumber(rawScan, vName)
            
            ' Replace in-place immediately so ONLY the real part number is displayed!
            modQCLogic.UnprotectConsole
            Me.Range("C12").Value = cleanPart
            
            ' Validate if part exists on active DA lines
            Dim partValid As Boolean
            partValid = modQCLogic.IsPartOnDA(daNum, cleanPart)
            
            If Not partValid Then
                Beep
                Me.Range("B3:J4").Value = "HOLD — PART [" & cleanPart & "] NOT FOUND IN ACTIVE DA!"
                Me.Range("B3:J4").Interior.Color = RGB(220, 38, 38)
                Me.Range("C12").Value = ""
                Me.Range("G12").Value = ""
                modQCLogic.ProtectConsole
                Me.Range("C12").Select
                GoTo CleanExit
            End If
            
            ' Part is valid! Check if SPQ is mentioned in the DA lines
            itemSPQ = modQCLogic.GetItemSPQ(daNum, cleanPart)
            
            If itemSPQ > 0 Then
                ' ZERO-TOUCH VERIFICATION: DA contains SPQ, auto-apply and process immediately!
                Me.Range("G12").Value = itemSPQ
                modQCLogic.ProtectConsole
                modQCLogic.ProcessScanConsole
            Else
                ' SPQ NOT in DA: Auto-open Quantity Input Pop-up Form!
                modQCLogic.ProtectConsole
                modQCLogic.PromptQuantityForPart cleanPart
            End If
        End If
    End If
    
    ' If user manually entered quantity (G12) and component is filled, auto-process
    If Not Intersect(Target, Me.Range("G12")) Is Nothing Then
        If Len(Trim(Me.Range("G12").Value)) > 0 And Len(Trim(Me.Range("C12").Value)) > 0 Then
            modQCLogic.ProcessScanConsole
        End If
    End If

CleanExit:
    Application.EnableEvents = True
End Sub

