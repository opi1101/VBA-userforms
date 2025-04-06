VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLog 
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765
   OleObjectBlob   =   "frmLog.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Source: https://github.com/opi1101

Private mCancelled As Boolean

Private Sub UserForm_Initialize()
  Caption = ThisWorkbook.Name
  lblHeader = vbNullString
  txtLog.Text = vbNullString
End Sub

Property Get Cancelled() As Boolean
  Cancelled = mCancelled
End Property

Sub AddLog(ByVal Text As String, Optional LineSeparator As String = "---")
  With txtLog
    Select Case .TextLength
      Case 0
        .Text = Text
      Case Else
        .Text = .Text & vbNewLine & LineSeparator & vbNewLine & Text
    End Select
  End With
End Sub

Sub ShowOrHide(FormModal As FormShowConstants, Optional Header As String)
  If Header <> vbNullString Then lblHeader = Header
  Select Case txtLog.TextLength
    Case 0
      Hide
    Case Else
      Show FormModal
  End Select
End Sub

Private Sub btnOk_Click()
  Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Cancel = True
  mCancelled = (CloseMode <> VbQueryClose.vbFormCode)
  Hide
End Sub
