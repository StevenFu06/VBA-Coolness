VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2730
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mouse_state As Integer
Dim cur_active As Integer
Dim global_month As String
Dim global_year As String

Private Sub SetVal(butt_num)
    set_date = global_year & "-" & global_month & "-" & Me.Controls("label" & butt_num).Caption
    Range("A1").Value = set_date
End Sub

Private Sub UserForm_Initialize()
    mouse_state = 0
    global_year = year(Date)
    global_month = Format(Date, "m")
    CalCaption
End Sub

Sub CalCaption()
    start_date = Weekday((global_year) & "-" & global_month & "-" & "1")
    num_days = MonthDays(global_month)
    previous_count = MonthDays(global_month - 1) - start_date + 2
    current_count = 1
    next_count = 1
    Month.Caption = MonthName(global_month) & " " & global_year
    For i = 1 To 42
        If i < start_date Then
            Me.Controls("label" & i).Caption = previous_count
            Me.Controls("Label" & i).BackColor = RGB(255, 255, 255)
            Me.Controls("Label" & i).ForeColor = RGB(120, 120, 120)
            previous_count = previous_count + 1
            
        ElseIf i >= start_date And current_count <= num_days Then
            Me.Controls("label" & i).Caption = current_count
            Me.Controls("Label" & i).BackColor = RGB(242, 242, 242)
            Me.Controls("Label" & i).ForeColor = RGB(0, 0, 0)
            
            current_count = current_count + 1
        ElseIf i > start_date And current_count > num_days Then
            Me.Controls("label" & i).Caption = next_count
            Me.Controls("Label" & i).BackColor = RGB(255, 255, 255)
            Me.Controls("Label" & i).ForeColor = RGB(120, 120, 120)
            next_count = next_count + 1
        End If
    Next i
End Sub

Private Function MonthDays(myMonth) As Integer
    MonthDays = Day(DateSerial(year(Date), myMonth + 1, 1) - 1)
End Function

Private Sub up_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    str_date = global_year & "-" & global_month & "-" & "1"
    str_date = DateAdd("m", 1, str_date)
    global_month = Format(str_date, "m")
    global_year = year(str_date)
    CalCaption
End Sub

Private Sub down_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    str_date = global_year & "-" & global_month & "-" & "1"
    str_date = DateAdd("m", -1, str_date)
    global_month = Format(str_date, "m")
    global_year = year(str_date)
    CalCaption
End Sub

'############################################################################################################################
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetInactive
End Sub

Private Sub SetInactive()
    For i = 1 To 42
        If i <> cur_active And Me.Controls("Label" & i).BorderColor <> RGB(255, 255, 255) Then
            Me.Controls("Label" & i).BackColor = Me.Controls("Label" & i).BackColor
            Me.Controls("Label" & i).BorderColor = RGB(255, 255, 255)
        End If
    Next i
End Sub

Private Sub MouseDown(butt_num)
    If Me.Controls("Label" & butt_num).BackColor <> RGB(255, 255, 255) Then
        mouse_state = 1
        cur_active = butt_num
        Me.Controls("Label" & butt_num).BackColor = Me.Controls("Label" & butt_num).BackColor
        Me.Controls("Label" & butt_num).BorderColor = RGB(0, 120, 215)
        SetVal butt_num
        SetInactive
    End If
End Sub

Private Sub MouseMove(butt_num)
    SetInactive
    If Me.Controls("Label" & butt_num).BorderColor <> RGB(200, 200, 200) And mouse_state = 0 And cur_active <> butt_num Then
        Me.Controls("Label" & butt_num).BackColor = Me.Controls("Label" & butt_num).BackColor
        Me.Controls("Label" & butt_num).BorderColor = RGB(200, 200, 200)
    End If
End Sub


'############################################################################################################################
Private Sub Label1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 1
End Sub
Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 1
End Sub
'############################################################################################################################
Private Sub Label2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 2
End Sub
Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 2
End Sub
'############################################################################################################################
Private Sub Label3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 3
End Sub
Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 3
End Sub
'############################################################################################################################
Private Sub Label4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 4
End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 4
End Sub
'############################################################################################################################
Private Sub Label5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 5
End Sub
Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 5
End Sub
'############################################################################################################################
Private Sub Label6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 6
End Sub
Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 6
End Sub
'############################################################################################################################
Private Sub Label7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 7
End Sub
Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 7
End Sub
'############################################################################################################################
Private Sub Label8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 8
End Sub
Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 8
End Sub
'############################################################################################################################
Private Sub Label9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 9
End Sub
Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 9
End Sub
'############################################################################################################################
Private Sub Label10_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label10_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 10
End Sub
Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 10
End Sub
'############################################################################################################################
Private Sub Label11_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 11
End Sub
Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 11
End Sub
'############################################################################################################################
Private Sub Label12_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 12
End Sub
Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 12
End Sub
'############################################################################################################################
Private Sub Label13_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 13
End Sub
Private Sub Label13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 13
End Sub
'############################################################################################################################
Private Sub Label14_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 14
End Sub
Private Sub Label14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 14
End Sub
'############################################################################################################################
Private Sub Label15_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label15_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 15
End Sub
Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 15
End Sub
'############################################################################################################################
Private Sub Label16_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label16_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 16
End Sub
Private Sub Label16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 16
End Sub
'############################################################################################################################
Private Sub Label17_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label17_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 17
End Sub
Private Sub Label17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 17
End Sub
'############################################################################################################################
Private Sub Label18_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label18_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 18
End Sub
Private Sub Label18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 18
End Sub
'############################################################################################################################
Private Sub Label19_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label19_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 19
End Sub
Private Sub Label19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 19
End Sub
'############################################################################################################################
Private Sub Label20_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label20_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 20
End Sub
Private Sub Label20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 20
End Sub
'############################################################################################################################
Private Sub Label21_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label21_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 21
End Sub
Private Sub Label21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 21
End Sub
'############################################################################################################################
Private Sub Label22_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label22_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 22
End Sub
Private Sub Label22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 22
End Sub
'############################################################################################################################
Private Sub Label23_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label23_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 23
End Sub
Private Sub Label23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 23
End Sub
'############################################################################################################################
Private Sub Label24_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label24_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 24
End Sub
Private Sub Label24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 24
End Sub
'############################################################################################################################
Private Sub Label25_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label25_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 25
End Sub
Private Sub Label25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 25
End Sub
'############################################################################################################################
Private Sub Label26_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label26_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 26
End Sub
Private Sub Label26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 26
End Sub
'############################################################################################################################
Private Sub Label27_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label27_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 27
End Sub
Private Sub Label27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 27
End Sub
'############################################################################################################################
Private Sub Label28_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label28_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 28
End Sub
Private Sub Label28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 28
End Sub
'############################################################################################################################
Private Sub Label29_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label29_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 29
End Sub
Private Sub Label29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 29
End Sub
'############################################################################################################################
Private Sub Label30_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label30_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 30
End Sub
Private Sub Label30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 30
End Sub
'############################################################################################################################
Private Sub Label31_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label31_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 31
End Sub
Private Sub Label31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 31
End Sub
'############################################################################################################################
Private Sub Label32_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label32_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 32
End Sub
Private Sub Label32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 32
End Sub
'############################################################################################################################
Private Sub Label33_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label33_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 33
End Sub
Private Sub Label33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 33
End Sub
'############################################################################################################################
Private Sub Label34_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label34_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 34
End Sub
Private Sub Label34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 34
End Sub
'############################################################################################################################
Private Sub Label35_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label35_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 35
End Sub
Private Sub Label35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 35
End Sub
'############################################################################################################################
Private Sub Label36_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label36_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 36
End Sub
Private Sub Label36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 36
End Sub
'############################################################################################################################
Private Sub Label37_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label37_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 37
End Sub
Private Sub Label37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 37
End Sub
'############################################################################################################################
Private Sub Label38_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label38_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 38
End Sub
Private Sub Label38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 38
End Sub
'############################################################################################################################
Private Sub Label39_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label39_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 39
End Sub
Private Sub Label39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 39
End Sub

'############################################################################################################################
Private Sub Label40_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label40_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 40
End Sub
Private Sub Label40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 40
End Sub
'############################################################################################################################
Private Sub Label41_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label41_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 41
End Sub
Private Sub Label41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 41
End Sub
'############################################################################################################################
Private Sub Label42_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mouse_state = 0
End Sub
Private Sub Label42_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseDown 42
End Sub
Private Sub Label42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMove 42
End Sub
