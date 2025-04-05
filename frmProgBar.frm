VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgBar 
   ClientHeight    =   1005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "frmProgBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
  Caption = ThisWorkbook.Name
  lblBar.Width = 0
  lblPerc = "0%"
  lblProgName = vbNullString
  lblCount = vbNullString
End Sub

Sub UpdateProgress(Current As Variant, Total As Variant, ProgName As String)
Dim p As Double

  On Error Resume Next
  p = Current / Total
  On Error GoTo 0
  
  Select Case True
    Case p > 1
      p = 1
    Case p < 0
      p = 0
  End Select
  lblBar.Width = lblBarScale.Width * p
  lblPerc = Int(p * 100) & "%"
  lblProgName = ProgName
  lblCount = Round(Current, 2) & " / " & Round(Total, 2)
  If Visible = False Then Show vbModeless
  DoEvents
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Cancel = (CloseMode <> VbQueryClose.vbFormCode)
End Sub
