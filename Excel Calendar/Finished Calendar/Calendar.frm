VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "Calendar"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2775
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str_month As String
Dim str_year As String

Private Sub OutputDate(val)
    str_day = Calendar.Controls("ToggleButton" & val).Caption
    str_date = str_year & "-" & str_month & "-" & str_day
    
'   Change line below to output
    ActiveSheet.Range("A1").value = str_date
End Sub

Private Sub UserForm_Initialize()
    str_month = month(Date)
    str_year = year(Date)
    loadMonth str_month, str_year
End Sub

Private Sub loadnum(start_button, num_days)
    count = 1
    For i = 1 To 42
        If i >= start_button And count <= num_days Then
            Calendar.Controls("ToggleButton" & i).Caption = count
            Calendar.Controls("ToggleButton" & i).Visible = True
            count = count + 1
        Else
            Calendar.Controls("ToggleButton" & i).Caption = ""
            Calendar.Controls("ToggleButton" & i).Visible = False
        End If
    Next i
End Sub

Private Sub calSelect(button)
    For i = 1 To 42
        If i <> button Then
            Calendar.Controls("ToggleButton" & i).value = False
        End If
    Next i
    OutputDate (button)
End Sub

Private Sub SpinButton1_SpinUp()
    direction = 1
    str_date = monthShift(str_month, str_year, direction)
    str_month = month(str_date)
    str_year = year(str_date)
    loadMonth str_month, str_year
End Sub

Private Sub SpinButton1_SpinDown()
    direction = 0
    str_date = monthShift(str_month, str_year, direction)
    str_month = month(str_date)
    str_year = year(str_date)
    loadMonth str_month, str_year
    Label8.Caption = MonthName(str_month) & " " & str_year
End Sub

Private Sub loadMonth(month, year)
    str_date = year & "-" & month & "-" & "1"
    month_len = MonthDays(month)
    loadnum Weekday(str_date), month_len
    Label8.Caption = MonthName(str_month) & " " & str_year
End Sub

Private Function MonthDays(myMonth) As Integer
    MonthDays = Day(DateSerial(year(Date), myMonth + 1, 1) - 1)
End Function

Private Function monthShift(month, year, direction) As String
    str_date = year & "-" & month & "-" & "1"
    If direction = 1 Then
        monthShift = DateAdd("m", 1, str_date)
    End If
    If direction = 0 Then
        monthShift = DateAdd("m", -1, str_date)
    End If
End Function

Private Sub ToggleButton1_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (1)
    OutputDate (1)
End Sub
Private Sub ToggleButton2_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (2)
End Sub
Private Sub ToggleButton3_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (3)
End Sub
Private Sub ToggleButton4_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (4)
End Sub
Private Sub ToggleButton5_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (5)
End Sub
Private Sub ToggleButton6_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (6)
End Sub
Private Sub ToggleButton7_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (7)
End Sub
Private Sub ToggleButton8_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (8)
End Sub
Private Sub ToggleButton9_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (9)
End Sub
Private Sub ToggleButton10_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (10)
End Sub
Private Sub ToggleButton11_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (12)
End Sub
Private Sub ToggleButton12_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (12)
End Sub
Private Sub ToggleButton13_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (13)
End Sub
Private Sub ToggleButton14_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (14)
End Sub
Private Sub ToggleButton15_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (15)
End Sub
Private Sub ToggleButton16_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (16)
End Sub
Private Sub ToggleButton17_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (17)
End Sub
Private Sub ToggleButton18_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (18)
End Sub
Private Sub ToggleButton19_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (19)
End Sub
Private Sub ToggleButton20_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (20)
End Sub
Private Sub ToggleButton21_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (21)
End Sub
Private Sub ToggleButton22_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (22)
End Sub
Private Sub ToggleButton23_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (23)
End Sub
Private Sub ToggleButton24_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (24)
End Sub
Private Sub ToggleButton25_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (25)
End Sub
Private Sub ToggleButton26_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (26)
End Sub
Private Sub ToggleButton27_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (27)
End Sub
Private Sub ToggleButton28_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (28)
End Sub
Private Sub ToggleButton29_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (29)
End Sub
Private Sub ToggleButton30_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (30)
End Sub
Private Sub ToggleButton31_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (31)
End Sub
Private Sub ToggleButton32_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (32)
End Sub
Private Sub ToggleButton33_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (33)
End Sub
Private Sub ToggleButton34_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (34)
End Sub
Private Sub ToggleButton35_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (35)
End Sub
Private Sub ToggleButton36_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (36)
End Sub
Private Sub ToggleButton37_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (37)
End Sub
Private Sub ToggleButton38_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (38)
End Sub
Private Sub ToggleButton39_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (39)
End Sub
Private Sub ToggleButton40_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (40)
End Sub
Private Sub ToggleButton41_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (41)
End Sub
Private Sub ToggleButton42_MouseDown(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    calSelect (42)
End Sub
