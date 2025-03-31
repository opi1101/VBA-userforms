VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "DatePicker"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmDatePicker.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mCancelled As Boolean
Private mDate As Double

Private Sub UserForm_Initialize()
  fillMonths
  fillDayNames
  DateToUserForm DateTime.Date
End Sub

Property Get Cancelled() As Boolean
  Cancelled = mCancelled
End Property

Property Get DateAsDouble() As Double
  DateAsDouble = mDate
End Property

Property Get SelectedYear() As Integer
On Error Resume Next
  SelectedYear = txtYear.Text
End Property

Property Get SelectedMonth() As Integer
On Error Resume Next
  SelectedMonth = (cboMonth.ListIndex + 1)
End Property

Sub DateToUserForm(Dat As Double)
  mDate = Dat
  txtYear.Text = Year(Dat)
  cboMonth.ListIndex = Month(Dat) - 1
  updateCalendar Dat
End Sub

Private Sub lblNext_Click()
  DateToUserForm DateSerial(SelectedYear, SelectedMonth + 1, 1)
End Sub

Private Sub lblPrev_Click()
  DateToUserForm DateSerial(SelectedYear, SelectedMonth - 1, 1)
End Sub

Private Sub lblToday_Click()
  DateToUserForm Date
End Sub

Private Sub lblTomorrow_Click()
  DateToUserForm DateAdd("d", 1, Date)
End Sub

Private Sub lblYesterday_Click()
  DateToUserForm DateAdd("d", -1, Date)
End Sub

Private Sub updateCalendar(ByVal Dat As Double)
Dim x As Byte, d As Byte, l As Byte, f As Byte

  l = Day(DateSerial(Year(Dat), Month(Dat) + 1, 1) - 1) ' Last day
  f = Weekday(DateSerial(Year(Dat), Month(Dat), 1), vbMonday) ' First weekday number
  For x = 3 To 44 ' Labels
    With Controls("Label" & x)
      Select Case True
        Case x - 2 < f, d >= l
          .Visible = False
        Case Else
          d = d + 1
          .Caption = d
          .Visible = True
          If d = Day(Dat) Then
            selectDay Controls("Label" & x)
          End If
      End Select
    End With
  Next x
End Sub

Private Sub cboMonth_Change()
Dim b As Boolean

  b = (cboMonth.ListIndex > -1)
  frCalendar.Visible = b
  If b = False Then
    MsgBox "Invalid month: " & cboMonth.Text & vbNewLine & "Please select month from the dropdown list.", vbOKOnly, Caption
    Exit Sub
  End If
  DateToUserForm DateSerial(SelectedYear, SelectedMonth, 1)
End Sub

Private Sub txtYear_Change()
Dim b As Boolean

  If txtYear.TextLength <> 4 Then Exit Sub
  b = (txtYear.Text Like "####")
  If b Then b = txtYear.Text >= Year(0)
  frCalendar.Visible = b
  If b = False Then
    MsgBox "Invalid year: " & txtYear.Text & vbNewLine & "Please provide year in 'yyyy' format.", vbOKOnly, Caption
    Exit Sub
  End If
  DateToUserForm DateSerial(SelectedYear, SelectedMonth, 1)
End Sub

Private Sub fillMonths()
Dim x As Byte

  With cboMonth
    .Clear
    For x = 1 To 12
      .AddItem StrConv(MonthName(x), vbProperCase)
    Next x
  End With
End Sub

Private Sub fillDayNames()
  lblMon = WeekdayName(1, True, vbMonday)
  lblTue = WeekdayName(2, True, vbMonday)
  lblWed = WeekdayName(3, True, vbMonday)
  lblThu = WeekdayName(4, True, vbMonday)
  lblFri = WeekdayName(5, True, vbMonday)
  lblSat = WeekdayName(6, True, vbMonday)
  lblSun = WeekdayName(7, True, vbMonday)
End Sub

Private Sub selectDay(Lbl As Object)
Dim x As Byte

  mDate = DateSerial(SelectedYear, SelectedMonth, Int(Lbl.Caption))
  For x = 3 To 44 ' Labels
    With Controls("Label" & x)
      .BackStyle = fmBackStyleTransparent
      .ForeColor = &H8000000D
      If .Name = Lbl.Name Then
        .BackStyle = fmBackStyleOpaque
        .ForeColor = vbWhite
      End If
    End With
  Next x
End Sub

Private Sub Label10_Click()
  selectDay Label10
End Sub

Private Sub Label11_Click()
  selectDay Label11
End Sub

Private Sub Label12_Click()
  selectDay Label12
End Sub

Private Sub Label13_Click()
  selectDay Label13
End Sub

Private Sub Label14_Click()
  selectDay Label14
End Sub

Private Sub Label15_Click()
  selectDay Label15
End Sub

Private Sub Label16_Click()
  selectDay Label16
End Sub

Private Sub Label17_Click()
  selectDay Label17
End Sub

Private Sub Label18_Click()
  selectDay Label18
End Sub

Private Sub Label19_Click()
  selectDay Label19
End Sub

Private Sub Label20_Click()
  selectDay Label20
End Sub

Private Sub Label21_Click()
  selectDay Label21
End Sub

Private Sub Label22_Click()
  selectDay Label22
End Sub

Private Sub Label23_Click()
  selectDay Label23
End Sub

Private Sub Label24_Click()
  selectDay Label24
End Sub

Private Sub Label25_Click()
  selectDay Label25
End Sub

Private Sub Label26_Click()
  selectDay Label26
End Sub

Private Sub Label27_Click()
  selectDay Label27
End Sub

Private Sub Label28_Click()
  selectDay Label28
End Sub

Private Sub Label29_Click()
  selectDay Label29
End Sub

Private Sub Label3_Click()
  selectDay Label3
End Sub

Private Sub Label30_Click()
  selectDay Label30
End Sub

Private Sub Label31_Click()
  selectDay Label31
End Sub

Private Sub Label32_Click()
  selectDay Label32
End Sub

Private Sub Label33_Click()
  selectDay Label33
End Sub

Private Sub Label34_Click()
  selectDay Label34
End Sub

Private Sub Label35_Click()
  selectDay Label35
End Sub

Private Sub Label36_Click()
  selectDay Label36
End Sub

Private Sub Label37_Click()
  selectDay Label37
End Sub

Private Sub Label38_Click()
  selectDay Label38
End Sub

Private Sub Label39_Click()
  selectDay Label39
End Sub

Private Sub Label4_Click()
  selectDay Label4
End Sub

Private Sub Label40_Click()
  selectDay Label40
End Sub

Private Sub Label41_Click()
  selectDay Label41
End Sub

Private Sub Label42_Click()
  selectDay Label42
End Sub

Private Sub Label43_Click()
  selectDay Label43
End Sub

Private Sub Label44_Click()
  selectDay Label44
End Sub

Private Sub Label5_Click()
  selectDay Label5
End Sub

Private Sub Label6_Click()
  selectDay Label6
End Sub

Private Sub Label7_Click()
  selectDay Label7
End Sub

Private Sub Label8_Click()
  selectDay Label8
End Sub

Private Sub Label9_Click()
  selectDay Label9
End Sub

Private Sub btnOk_Click()
  Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Cancel = True
  mCancelled = (CloseMode <> VbQueryClose.vbFormCode)
  Hide
End Sub
