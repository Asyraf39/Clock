VERSION 5.00
Begin VB.Form Clock 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSaveL 
      Caption         =   "Save lap time"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveS 
      Caption         =   "Save split time"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   27
      Top             =   6720
      Width           =   975
   End
   Begin VB.CheckBox chkLink 
      BackColor       =   &H80000016&
      Caption         =   "Check1"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   6840
      Width           =   255
   End
   Begin VB.ListBox lstLap 
      Height          =   1815
      ItemData        =   "Clock (Advanced).frx":0000
      Left            =   6480
      List            =   "Clock (Advanced).frx":0002
      TabIndex        =   22
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ListBox lstSplit 
      Height          =   1815
      ItemData        =   "Clock (Advanced).frx":0004
      Left            =   6480
      List            =   "Clock (Advanced).frx":0006
      TabIndex        =   21
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdLap 
      Caption         =   "Lap"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5160
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3960
      TabIndex        =   19
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdStartS 
      Caption         =   "Start"
      Height          =   615
      Left            =   3960
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdPauseS 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5160
      TabIndex        =   17
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdStopS 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3960
      TabIndex        =   16
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdResetS 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5160
      TabIndex        =   15
      Top             =   5880
      Width           =   975
   End
   Begin VB.Timer Stopwatch 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   2160
   End
   Begin VB.CommandButton cmdResetT 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdStopT 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdPauseT 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      TabIndex        =   9
      Top             =   5160
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   2160
   End
   Begin VB.CommandButton cmdStartT 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtSecond 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   6
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtMinute 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Width           =   735
   End
   Begin VB.Timer Clock 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblLink 
      BackColor       =   &H80000016&
      Caption         =   "Link with stopwatch"
      Height          =   255
      Left            =   840
      TabIndex        =   26
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      Caption         =   "Split"
      Height          =   255
      Left            =   6600
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblLap 
      Alignment       =   2  'Center
      Caption         =   "Lap"
      Height          =   255
      Left            =   6600
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblTimeElapsed 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00.000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   14
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblStopwatch 
      Alignment       =   2  'Center
      Caption         =   "Stopwatch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3360
      Y1              =   2160
      Y2              =   7440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblTimeLeft 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00.000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Clock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblClock 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Timer
Dim a1 As Integer 'Minute
Dim a2 As String 'Minute for display
Dim a3 As Integer 'Time left in minutes
Dim b1 As Integer 'Second
Dim b2 As String 'Second for display
Dim b3 As Integer 'Time left in seconds
Dim c1 As Integer 'Millisecond
Dim c2 As String 'Millisecond for display
Dim c3 As Integer 'Time left in milliseconds
Dim d As Integer 'Total time in seconds
Dim e1 As Double 'Time elapsed
Dim e2 As Integer 'Time left as whole number
Dim e3 As Double 'Time left
Dim f1 As Double 'Start time
Dim f2 As Double 'Current time

'Stopwatch
Dim g1 As Integer 'Minute
Dim g2 As String 'Minute for display
Dim g3 As Integer 'Time elapsed in minutes
Dim h1 As Integer 'Second
Dim h2 As String 'Second for display
Dim h3 As Integer 'Time elapsed in seconds
Dim i1 As Integer 'Millisecond
Dim i2 As String 'Millisecond for display
Dim i3 As Integer 'Time elapsed in milliseconds
Dim j1 As Double 'Start time
Dim j2 As Double 'Current time
Dim k1 As Double 'Time elapsed
Dim k2 As Integer 'Time elapsed as whole number
Dim l1 As Double 'Total time in seconds
Dim l2 As Double 'Total time in seconds for timer
Dim m1 As Integer 'Split
Dim m2 As String 'Split for display
Dim n1 As Integer 'Lap
Dim n2 As String 'Lap for display

Private Sub chkLink_Click()

If chkLink.Value = 1 Then 'If link is checked
    If lstSplit.List(0) <> "" Or lstLap.List(0) <> "" Then 'If one of the listbox is filled
        If Timer1.Enabled = False Then 'If timer is not running
            If MsgBox("This will clear your records!" & vbCrLf & "Continue?", vbYesNo + vbExclamation + vbDefaultButton2, "Records available") = vbYes Then 'If user clicks yes
                cmdStartS.Enabled = False 'Disable start button
                cmdSaveS.Enabled = False 'Disable save split time button
                cmdSaveL.Enabled = False 'Disable save lap time button
                cmdResetS.Enabled = False 'Disable reset button
                lblTimeElapsed.Caption = "00:00.000" 'Reset display
                m1 = 0 'Reset split
                lstSplit.Clear 'Clear split
                n1 = 0 'Reset lap
                lstLap.Clear 'Clear lap
            Else 'If user clicks no
            chkLink.Value = 0 'Uncheck link
            End If
        Else 'If timer is running
            cmdStartS.Enabled = False 'Disable start button
            If e3 > 0 Then
                m1 = 0 'Reset split
                lstSplit.Clear 'Clear split
                n1 = 0 'Reset lap
                lstLap.Clear 'Clear lap
            End If
        End If
    Else 'If both listbox is empty
        cmdStartS.Enabled = False 'Disable start button
        cmdSaveS.Enabled = False 'Disable save split time button
        cmdSaveL.Enabled = False 'Disable save lap time button
        cmdResetS.Enabled = False 'Disable reset button
        lblTimeElapsed.Caption = "00:00.000" 'Reset display
        m1 = 0 'Reset split
        lstSplit.Clear 'Clear split
        n1 = 0 'Reset lap
        lstLap.Clear 'Clear lap
    End If
    lblTimeElapsed.BackColor = lblTimeLeft.BackColor 'Change display color
    If cmdStartT.Caption = "Resume" Then 'If resume
        l2 = (g1 * 60) + h1 + (i1 / 1000) 'Assign total time in seconds for timer
        k2 = Fix(e1 + l2) 'Assign time elapsed as whole number
        i3 = ((e1 + l2) - k2) * 1000 'Assign time elapsed in milliseconds
        g3 = Fix(k2 / 60) 'Assign time elapsed in minutes
        h3 = k2 Mod 60 'Assign time elapsed in seconds
        If g3 < 10 Then 'If time elapsed in minutes is less than 10
            g2 = "0" & g3 'Fill in 0 before minute for display
        Else 'If time elapsed in minutes is not less than 10
            g2 = g3 'Assign minute for display
        End If
        If h3 < 10 Then 'If time elapsed in seconds is less than 10
            h2 = "0" & h3 'Fill in 0 before second for display
        Else 'If time elapsed in seconds is not less than 10
            h2 = h3 'Assign second for display
        End If
        If i3 < 10 Then 'If time elapsed in milliseconds is less than 10
            i2 = "00" & i3 'Fill in 0 before millisecond for display
        ElseIf i3 < 100 Then 'If time elapsed in milliseconds is less than 100
            i2 = "0" & i3 'Fill in 0 before millisecond for display
        Else 'If time elapsed in minutes is not less than 100
            i2 = i3 'Assign millisecond for display
        End If
        lblTimeElapsed.Caption = g2 & ":" & h2 & "." & i2 'Display stopwatch
    End If
    If lblTimeElapsed <> "00:00.000" Then 'If stopwatch is running
        cmdSplit.Enabled = True 'Enable split button
        cmdLap.Enabled = True 'Enable lap button
    End If
ElseIf chkLink.Value = 0 Then 'If link is unchecked
    lblTimeElapsed.Caption = "00:00.000" 'Reset display
    cmdStartS.Enabled = True 'Enable start button
    cmdSplit.Enabled = False 'Enable split button
    cmdLap.Enabled = False 'Enable lap button
    If lstSplit.List(0) <> "" Or lstLap.List(0) <> "" Then 'If one of the listbox is filled
        cmdResetS.Enabled = True 'Enable reset button
    End If
    lblTimeElapsed.BackColor = vbButtonFace 'Revert display color
End If

End Sub

Private Sub cmdLap_Click()

n1 = n1 + 1 'Add 1 to lap
If n1 < 10 Then 'If lap is less than 10
    n2 = "0" & n1 'Fill in 0 before lap for display
Else 'If lap is not less than 10
    n2 = n1 'Assign lap for display
End If
If m1 <> 0 Then 'If split is not 0
    lstSplit.AddItem "Lap " & n2, lstSplit.ListCount - m1 'Insert lap to split
    lstSplit.AddItem "" 'Insert blank space
    lstSplit.ListIndex = lstSplit.ListCount - 1 'Show last line
    lstSplit.ListIndex = lstSplit.ListCount - m1 - 2 'Show lap
    m1 = 0 'Reset split
End If
lstLap.AddItem n2 & ".  " & lblTimeElapsed.Caption 'Insert time to lap
lstLap.ListIndex = lstLap.ListCount - 1 'Show last line
If n1 = 99 Then 'If lap is 99
    cmdLap.Enabled = False 'Disable lap button
End If
g3 = 0 'Assign time elapsed in minutes as 0
h3 = 0 'Assign time elapsed in seconds as 0
i3 = 0 'Assign time elapsed in milliseconds as 0
j1 = Timer 'Set start time

End Sub

Private Sub cmdPauseS_Click()

If lstSplit.List(0) <> "" Then 'If split time is filled
    cmdSaveS.Enabled = True 'Enable save split time button
End If
If lstLap.List(0) <> "" Then 'If lap time is filled
    cmdSaveL.Enabled = True 'Enable save lap time button
End If
Stopwatch.Enabled = False 'Pause stopwatch
cmdPauseS.Enabled = False 'Disable pause button
cmdStartS.Enabled = True 'Enable resume button
cmdStopS.Enabled = True 'Enable stop button

End Sub

Private Sub cmdPauseT_Click()

If lstSplit.List(0) <> "" Then 'If split time is filled
    cmdSaveS.Enabled = True 'Enable save split time button
End If
If lstLap.List(0) <> "" Then 'If lap time is filled
    cmdSaveL.Enabled = True 'Enable save lap time button
End If
Timer1.Enabled = False 'Pause timer
cmdStartT.Enabled = True 'Enable resume button
cmdStopT.Enabled = True 'Enable stop button

End Sub

Private Sub cmdResetS_Click()

cmdSaveS.Enabled = False 'Enable save split time button
cmdSaveL.Enabled = False 'Enable save lap time button
cmdResetS.Enabled = False 'Disable reset button
lblTimeElapsed.Caption = "00:00.000" 'Reset display
m1 = 0 'Reset split
lstSplit.Clear 'Clear split
n1 = 0 'Reset lap
lstLap.Clear 'Clear lap

End Sub

Private Sub cmdResetT_Click()

txtMinute.Text = "" 'Clear minute
txtSecond.Text = "" 'Clear second
lblTimeElapsed.Caption = "00:00.000" 'Reset display

End Sub

Private Sub cmdSaveL_Click()

'Program directory
Dim file As String 'Declare directory
file = App.Path & "\Lap.txt" 'Assign savefile directory

If Dir(file) <> "" Then 'If savefile exists
    If MsgBox("Save data?" & vbCrLf & "(This will overwrite previous savefile)", vbYesNo + vbQuestion, "Save") = vbYes Then 'If user clicks yes
        GoTo Code: 'Run program code
    End If
Else 'If savefile does not exists
    If MsgBox("Save data?", vbYesNo + vbQuestion, "Save") = vbYes Then 'If user clicks yes
        GoTo Code: 'Run program code
    End If
End If

Exit Sub

Code: 'Program code
save = FreeFile 'Assign savefile with a number
Open file For Output As #save 'Open savefile
For laptime = 0 To lstLap.ListCount - 1 'For every lap time
    Print #save, lstLap.List(laptime) 'Save lap time
Next 'Repeat save
Close #1 'Close savefile

End Sub

Private Sub cmdSaveS_Click()

'Program directory
Dim file As String 'Declare directory
file = App.Path & "\Split.txt" 'Assign savefile directory

If Dir(file) <> "" Then 'If savefile exists
    If MsgBox("Save data?" & vbCrLf & "(This will overwrite previous savefile)", vbYesNo + vbQuestion, "Save") = vbYes Then 'If user clicks yes
        GoTo Code: 'Run program code
    End If
Else 'If savefile does not exists
    If MsgBox("Save data?", vbYesNo + vbQuestion, "Save") = vbYes Then 'If user clicks yes
        GoTo Code: 'Run program code
    End If
End If

Exit Sub

Code: 'Program code
save = FreeFile 'Assign savefile with a number
Open file For Output As #save 'Open savefile
For splittime = 0 To lstSplit.ListCount - 1 'For every split time
    Print #save, lstSplit.List(splittime) 'Save split time
Next 'Repeat save
Close #1 'Close savefile

End Sub

Private Sub cmdSplit_Click()

m1 = m1 + 1 'Add 1 to split
If m1 < 10 Then 'If split is less than 10
    m2 = "0" & m1 'Fill in 0 before split for display
Else 'If split is not less than 10
    m2 = m1 'Assign split for display
End If
lstSplit.AddItem m2 & ".  " & lblTimeElapsed.Caption 'Insert time to split
lstSplit.ListIndex = lstSplit.ListCount - 1 'Show last line
If m1 = 99 Then 'If split is 99
    cmdSplit.Enabled = False 'Disable split button
End If

End Sub

Private Sub cmdStartS_Click()

If cmdStartS.Caption = "Start" Then 'If start
    If lstSplit.List(0) <> "" Or lstLap.List(0) <> "" Then 'If one of the listbox is filled
        If MsgBox("This will clear your records!" & vbCrLf & "Continue?", vbYesNo + vbExclamation + vbDefaultButton2, "Records available") = vbYes Then 'If user clicks yes
            lstSplit.Clear 'Clear split
            m1 = 0 'Reset split
            lstLap.Clear 'Clear lap
            n1 = 0 'Reset lap
            g3 = 0 'Assign time elapsed in minutes as 0
            h3 = 0 'Assign time elapsed in seconds as 0
            i3 = 0 'Assign time elapsed in milliseconds as 0
            chkLink.Enabled = False 'Disable link
            cmdSplit.Enabled = True 'Enable split button
            cmdLap.Enabled = True 'Enable lap button
            cmdSaveS.Enabled = False 'Enable save split time button
            cmdSaveL.Enabled = False 'Enable save lap time button
            cmdStartS.Enabled = False 'Disable start button
            cmdStopS.Enabled = False 'Disable stop button
            cmdResetS.Enabled = False 'Disable reset button
            cmdPauseS.Enabled = True 'Enable pause button
            cmdStartS.Caption = "Resume" 'Change start button to resume button
            j1 = Timer 'Set start time
            Stopwatch.Enabled = True 'Start stopwatch
        End If
    Else 'If both listbox are empty
        g3 = 0 'Assign time elapsed in minutes as 0
        h3 = 0 'Assign time elapsed in seconds as 0
        i3 = 0 'Assign time elapsed in milliseconds as 0
        chkLink.Enabled = False 'Disable link
        cmdSplit.Enabled = True 'Enable split button
        cmdLap.Enabled = True 'Enable lap button
        cmdStartS.Enabled = False 'Disable start button
        cmdStopS.Enabled = False 'Disable stop button
        cmdResetS.Enabled = False 'Disable reset button
        cmdPauseS.Enabled = True 'Enable pause button
        cmdStartS.Caption = "Resume" 'Change start button to resume button
        j1 = Timer 'Set start time
        Stopwatch.Enabled = True 'Start stopwatch
    End If
Else 'If resume
    g3 = g2 'Assign time elapsed in minutes from display
    h3 = h2 'Assign time elapsed in seconds from display
    i3 = i2 'Assign time elapsed in milliseconds from display
    cmdSaveS.Enabled = False 'Disable save split time button
    cmdSaveL.Enabled = False 'Disable save lap time button
    cmdStartS.Enabled = False 'Disable resume button
    cmdStopS.Enabled = False 'Disable stop button
    cmdPauseS.Enabled = True 'Enable start button
    j1 = Timer 'Set start time
    Stopwatch.Enabled = True 'Start stopwatch
End If

End Sub

Private Sub cmdStartT_Click()

If cmdStartT.Caption = "Start" Then 'If start
    cmdStartT.Caption = "Resume" 'Change start button to resume button
    If txtMinute.Text <> "" Then 'If minute is filled
        a1 = txtMinute.Text 'Assign minute
    Else 'If minute is empty
        a1 = 0 'Assign minute as 0
        txtMinute.Text = "0" 'Display minute as zero
    End If
    If a1 < 10 Then 'If minute is less than 10
        a2 = "0" & a1 'Insert 0 before minute for display
    Else 'If minute is not less than 10
        a2 = a1 'Assign minute for display
    End If
    If b1 = 0 Then 'If second is 0
        b2 = "00" 'Assign second as 0 for display
    End If
    If txtSecond.Text <> "" Then 'If second is filled
        b1 = txtSecond.Text 'Assign second
    Else 'If second is empty
        b1 = 0 'Assign second as 0
        txtSecond.Text = "0" 'Display second as zero
    End If
    If b1 < 10 Then 'If second is less than 10
        b2 = "0" & b1 'Insert 0 before second for display
    Else 'If second is not less than 10
        b2 = b1 'Assign second for display
    End If
    If a1 = 0 Then 'If minute is 0
        a2 = "00" 'Assign minute as 0 for display
    End If
    txtMinute.Enabled = False 'Disable input for minute
    txtSecond.Enabled = False 'Disable input for second
    If chkLink.Value = 1 Then 'If link is checked
        cmdSplit.Enabled = True 'Enable split button
        cmdLap.Enabled = True 'Enable lap button
    End If
    cmdPauseT.Enabled = True 'Disable pause button
    cmdStartT.Enabled = False 'Disable start button
    cmdResetT.Enabled = False 'Disable reset button
    f1 = Timer 'Assign start time
    Timer1.Enabled = True 'Start timer
Else 'If resume
    a1 = a3 'Assign minute from display
    b1 = b3 'Assign second from display
    c1 = c3 'Assign millisecond from display
    g1 = g3 'Assign minute from display
    h1 = h3 'Assign second from display
    i1 = i3 'Assign millisecond from display
    If chkLink.Value = 1 Then 'If link is checked
        cmdSaveS.Enabled = False 'Disable save split time button
        cmdSaveL.Enabled = False 'Disable save lap time button
    End If
    cmdStartT.Enabled = False 'Disable start button
    cmdStopT.Enabled = False 'Disable stop button
    cmdResetT.Enabled = False 'Disable reset button
    f1 = Timer 'Assign start time
    Timer1.Enabled = True 'Start timer
End If

End Sub

Private Sub cmdStopS_Click()

chkLink.Enabled = True 'Enable link
cmdSplit.Enabled = False 'Enable split button
cmdLap.Enabled = False 'Enable lap button
cmdStartS.Enabled = True 'Enable start button
cmdPauseS.Enabled = False 'Disable pause button
cmdStopS.Enabled = False 'Disable stop button
cmdResetS.Enabled = True 'Enable reset button
cmdStartS.Caption = "Start" 'Change resume button to start button
g1 = 0 'Assign minute as 0
h1 = 0 'Assign second as 0
i1 = 0 'Assign millisecond as 0

End Sub

Private Sub cmdStopT_Click()

If txtMinute.Text = "0" Then 'If minute is 0
    txtMinute.Text = "" 'Display minute as empty
End If
If txtSecond.Text = "0" Then 'If second is 0
    txtSecond.Text = "" 'Display second as empty
End If
g1 = 0 'Assign minute as 0
h1 = 0 'Assign second as 0
i1 = 0 'Assign millisecond as 0
Timer1.Enabled = False 'Stop timer
txtMinute.Enabled = True 'Enable input for minute
txtSecond.Enabled = True 'Enable input for second
cmdResetT.Enabled = True 'Enbable reset button
cmdPauseT.Enabled = False 'Disable pause button
cmdStopT.Enabled = False 'Disable stop button
cmdStartT.Caption = "Start" 'Change resume button to start button
lblTimeLeft.BackColor = vbButtonFace 'Revert display color

End Sub

Private Sub Clock_Timer()

lblClock.Caption = Time 'Display time
lblDate.Caption = Date 'Displat date

End Sub

Private Sub Stopwatch_Timer()

j2 = Timer 'Set current time
l1 = (g3 * 60) + h3 + (i3 / 1000) 'Assign total time in seconds
k1 = Round(j2 - j1, 3) + l1 'Assign time elapsed
If k1 < 6000 Then 'If time elapsed is less than 6000
    k2 = Fix(k1) 'Assign time elapsed as whole number
    i1 = (k1 - k2) * 1000 'Assign time elapsed in milliseconds
    g1 = Fix(k2 / 60) 'Assign time elapsed in minutes
    h1 = k2 Mod 60 'Assign time elapsed in seconds
    If g1 < 10 Then 'If time elapsed in minutes is less than 10
        g2 = "0" & g1 'Fill in 0 before minute for display
    Else 'If time elapsed in minutes is not less than 10
        g2 = g1 'Assign minute for display
    End If
    If h1 < 10 Then 'If time elapsed in seconds is less than 10
        h2 = "0" & h1 'Fill in 0 before second for display
    Else 'If time elapsed in seconds is not less than 10
        h2 = h1 'Assign second for display
    End If
    If i1 < 10 Then 'If time elapsed in milliseconds is less than 10
        i2 = "00" & a1 'Fill in 0 before millisecond for display
    ElseIf i1 < 100 Then 'If time elapsed in milliseconds is less than 100
        i2 = "0" & i1 'Fill in 0 before millisecond for display
    Else 'If time elapsed in minutes is not less than 100
        i2 = i1 'Assign millisecond for display
    End If
    lblTimeElapsed.Caption = g2 & ":" & h2 & "." & i2 'Display stopwatch
Else 'If time elapsed is not less than 6000
    Stopwatch.Enabled = False 'Stop timer
    lblTimeElapsed.Caption = "99:99.999" 'Display timer
    MsgBox "Stopwatch time limit reached!", vbExclamation, "Time limit reached" 'Inform user
    cmdSplit.Enabled = False 'Disable split button
    cmdLap.Enabled = False 'Disable lap button
    cmdStartS.Enabled = True 'Enable start button
    cmdPauseS.Enabled = False 'Disable pause button
    cmdStopS.Enabled = False 'Disable stop button
    cmdResetS.Enabled = True 'Enable reset button
    cmdStartS.Caption = "Start" 'Change resume button to start button
End If

End Sub

Private Sub Timer1_Timer()

f2 = Timer 'Set current time
e1 = f2 - f1 'Assign time elapsed
d = (a1 * 60) + b1 + (c1 / 1000) 'Assign total time in seconds
If chkLink.Value = 1 Then 'If link is checked
    l2 = (g1 * 60) + h1 + (i1 / 1000)
    k2 = Fix(e1 + l2) 'Assign time elapsed as whole number
    i3 = ((e1 + l2) - k2) * 1000 'Assign time elapsed in milliseconds
    g3 = Fix(k2 / 60) 'Assign time elapsed in minutes
    h3 = k2 Mod 60 'Assign time elapsed in seconds
    If g3 < 10 Then 'If time elapsed in minutes is less than 10
        g2 = "0" & g3 'Fill in 0 before minute for display
    Else 'If time elapsed in minutes is not less than 10
        g2 = g3 'Assign minute for display
    End If
    If h3 < 10 Then 'If time elapsed in seconds is less than 10
        h2 = "0" & h3 'Fill in 0 before second for display
    Else 'If time elapsed in seconds is not less than 10
        h2 = h3 'Assign second for display
    End If
    If i3 < 10 Then 'If time elapsed in milliseconds is less than 10
        i2 = "00" & i3 'Fill in 0 before millisecond for display
    ElseIf i3 < 100 Then 'If time elapsed in milliseconds is less than 100
        i2 = "0" & i3 'Fill in 0 before millisecond for display
    Else 'If time elapsed in minutes is not less than 100
        i2 = i3 'Assign millisecond for display
    End If
    lblTimeElapsed.Caption = g2 & ":" & h2 & "." & i2 'Display stopwatch
End If
e3 = Round(d - e1, 3) 'Assign time left
e2 = Fix(e3) 'Assign time left as whole number
c3 = (e3 - e2) * 1000 'Assign time left in milliseconds
a3 = Fix(e2 / 60) 'Assign time left in minutes
b3 = e2 Mod 60 'Assign time left in seconds
If a3 < 10 Then 'If time left in minutes is less than 10
    a2 = "0" & a3 'Fill in 0 before minute for display
Else 'If time left in minutes is not less than 10
    a2 = a3 'Assign minute for display
End If
If b3 < 10 Then 'If time left in seconds is less than 10
    b2 = "0" & b3 'Fill in 0 before second for display
Else 'If time left in seconds is not less than 10
    b2 = b3 'Assign second for display
End If
If c3 < 10 Then 'If time left in milliseconds is less than 10
    c2 = "00" & c3 'Fill in 0 before millisecond for display
ElseIf c3 < 100 Then 'If time left in milliseconds is less than 100
    c2 = "0" & c3 'Fill in 0 before millisecond for display
Else 'If time left in milliseconds is not less than 100
    c2 = c3 'Assign millisecond for display
End If
lblTimeLeft.Caption = a2 & ":" & b2 & "." & c2 'Display timer
If b3 <= 2 And a3 = 0 Then 'If 2 seconds left
    If c3 > 875 Then 'If time left in milliseconds is more than 875
        lblTimeLeft.BackColor = &HC0C0FF 'Change display color
    ElseIf c3 > 750 Then 'If time left in milliseconds is more than 750
        lblTimeLeft.BackColor = vbButtonFace 'Revert display color
    ElseIf c3 > 625 Then 'If time left in milliseconds is more than 625
        lblTimeLeft.BackColor = &HC0C0FF 'Change display color
    ElseIf c3 > 500 Then 'If time left in milliseconds is more than 500
        lblTimeLeft.BackColor = vbButtonFace 'Revert display color
    ElseIf c3 > 375 Then 'If time left in milliseconds is more than 375
        lblTimeLeft.BackColor = &HC0C0FF 'Change display color
    ElseIf c3 > 250 Then 'If time left in milliseconds is more than 250
        lblTimeLeft.BackColor = vbButtonFace 'Revert display color
    ElseIf c3 > 125 Then 'If time left in milliseconds is more than 125
        lblTimeLeft.BackColor = &HC0C0FF 'Change display color
    Else 'If time left in milliseconds is less than 125
        lblTimeLeft.BackColor = vbButtonFace 'Revert display color
    End If
ElseIf b3 <= 5 And a3 = 0 Then 'If 5 seconds left
    If c3 > 750 Then 'If time left in milliseconds is more than 750
        lblTimeLeft.BackColor = &HC0C0FF 'Change display color
    ElseIf c3 > 500 Then 'If time left in milliseconds is more than 500
        lblTimeLeft.BackColor = vbButtonFace 'Revert display color
    ElseIf c3 > 250 Then 'If time left in milliseconds is more than 250
        lblTimeLeft.BackColor = &HC0C0FF 'Change display color
    Else 'If time left in milliseconds is less than 250
        lblTimeLeft.BackColor = vbButtonFace 'Revert display color
    End If
ElseIf b3 <= 10 And a3 = 0 Then 'If 10 seconds left
    If c3 > 500 Then 'If time left in milliseconds is more than 500
        lblTimeLeft.BackColor = &HC0C0FF 'Change display color
    Else 'If time left in milliseconds is less than 500
        lblTimeLeft.BackColor = vbButtonFace 'Revert display color
    End If
End If
If chkLink.Value = 1 Then 'If link is checked
    lblTimeElapsed.BackColor = lblTimeLeft.BackColor 'Change display color
    cmdSplit.Enabled = True 'Enable split button
    cmdLap.Enabled = True 'Enable lap button
ElseIf chkLink.Value = 0 Then 'If link is unchecked
    lblTimeElapsed.BackColor = vbButtonFace 'Revert display color
End If
If e3 < 0 Then 'If time left is less than 0
    Timer1.Enabled = False 'Stop timer
    txtMinute.Enabled = True 'Enable input for minute
    txtSecond.Enabled = True 'Enable input for second
    cmdResetT.Enabled = True 'Enbable reset button
    cmdPauseT.Enabled = False 'Disable pause button
    cmdStopT.Enabled = False 'Disable stop button
    cmdStartT.Enabled = True 'Enable start button
    cmdStartT.Caption = "Start" 'Change resume button to start button
    lblTimeLeft.Caption = "00:00.000" 'Reset display
    If chkLink.Value = 1 Then 'If link is checked
        If a1 < 10 Then 'If time elapsed in minutes is less than 10
            g2 = "0" & a1 'Fill in 0 before minute for display
        Else 'If time elapsed in minutes is not less than 10
            g2 = a1 'Assign minute for display
        End If
        If b1 < 10 Then 'If time elapsed in seconds is less than 10
            h2 = "0" & b1 'Fill in 0 before second for display
        Else 'If time elapsed in seconds is not less than 10
            h2 = b1 'Assign second for display
        End If
        lblTimeElapsed.Caption = g2 & ":" & h2 & ".000" 'Display stopwatch
    End If
    MsgBox "Time's up!", vbExclamation, "Time's up" 'Inform user
    If txtMinute.Text = "0" Then 'If minute is 0
        txtMinute.Text = "" 'Display minute as empty
    End If
    If txtSecond.Text = "0" Then 'If second is 0
        txtSecond.Text = "" 'Display second as empty
    End If
    lblTimeLeft.Caption = "00:00.000" 'Reset display
    If chkLink.Value = 1 Then 'If link is checked
        lblTimeElapsed.Caption = g2 & ":" & h2 & ".000" 'Display stopwatch
    End If
    cmdSplit.Enabled = False 'Disable split button
    cmdLap.Enabled = False 'Disable lap button
End If

End Sub

Private Sub txtMinute_Change()

On Error GoTo ErrorHandler 'If there is an error, go to the error handler

Code: 'Program code
If txtMinute.Text <> "" Or txtSecond.Text <> "" Then 'If one of the fields are filled
    cmdStartT.Enabled = True 'Enable start button
    cmdResetT.Enabled = True 'Enable reset button
Else 'If both fields are empty
    cmdStartT.Enabled = False 'Disable start button
    cmdPauseT.Enabled = False 'Disable pause button
    cmdStopT.Enabled = False 'Disable stop button
    cmdResetT.Enabled = False 'Disable reset button
End If

If txtMinute.Text <> "" Then 'If minute is filled
    If txtMinute.Text = "." & a1 Then 'If minute has a front decimal point
        txtMinute.Text = a1 'Remove decimal point
    End If
    a1 = txtMinute.Text 'Assign minute
    txtMinute.Text = a1 'Remove unwanted zeros and decimals
    If a1 < 100 And a1 >= 0 Then 'If minute is valid
        If cmdStartT.Caption = "Start" Then 'If start
            If txtMinute.Text = "0" Then 'If minute is 0
                txtMinute.Text = "" 'Display minute as empty
            End If
        End If
    Else 'If minute is invalid
        MsgBox "Please insert a valid number! (1-99)", vbExclamation, "Invalid value" 'Inform user
        txtMinute.Text = "" 'Display minute as empty
    End If
Else 'If minute is empty
    a1 = 0 'Assign minute as 0
End If

If a1 < 10 Then 'If minute is less than 10
    a2 = "0" & a1 'Insert 0 before minute for display
Else 'If minute is not less than 10
    a2 = a1 'Assign minute for display
End If
If b1 = 0 Then 'If second is 0
    b2 = "00" 'Assign second as 0 for display
End If
If b1 < 10 Then 'If second is less than 10
    b2 = "0" & b1 'Insert 0 before second for display
Else 'If second is not less than 10
    b2 = b1 'Assign second for display
End If
If a1 = 0 Then 'If minute is 0
    a2 = "00" 'Assign minute as 0 for display
End If
lblTimeLeft.Caption = a2 & ":" & b2 & ".000" 'Display timer

Exit Sub

ErrorHandler: 'Handles error
If a1 = 0 Then 'If minute is 0
    txtMinute.Text = "" 'Display minute as empty
Else 'If minute is filled
    txtMinute.Text = a1 'Disable value change
End If
Resume Code: 'Run program code

End Sub

Private Sub txtSecond_Change()

On Error GoTo ErrorHandler 'If there is an error, go to the error handler

Code: 'Program code
If txtMinute.Text <> "" Or txtSecond.Text <> "" Then 'If one of the fields are filled
    cmdStartT.Enabled = True 'Enable start button
    cmdResetT.Enabled = True 'Enable reset button
Else 'If both fields are empty
    cmdStartT.Enabled = False 'Disable start button
    cmdPauseT.Enabled = False 'Disable pause button
    cmdStopT.Enabled = False 'Disable stop button
    cmdResetT.Enabled = False 'Disable reset button
End If

If txtSecond.Text <> "" Then 'If second is filled
    If txtSecond.Text = "." & b1 Then 'If second has a front decimal point
        txtSecond.Text = b1 'Remove decimal point
    End If
    b1 = txtSecond.Text 'Assign second
    txtSecond.Text = b1 'Remove unwanted zeros and decimals
    If b1 < 60 And b1 >= 0 Then 'If second is valid
        If cmdStartT.Caption = "Start" Then 'If start
            If txtSecond.Text = "0" Then 'If second is 0
                txtSecond.Text = "" 'Display second as empty
            End If
        End If
    Else 'If second is invalid
        MsgBox "Please insert a valid number! (1-59)", vbExclamation, "Invalid value" 'Inform user
        txtSecond.Text = "" 'Display second as empty
    End If
Else 'If second is empty
    b1 = 0 'Assign second as 0
End If

If a1 < 10 Then 'If minute is less than 10
    a2 = "0" & a1 'Insert 0 before minute for display
Else 'If minute is not less than 10
    a2 = a1 'Assign minute for display
End If
If b1 = 0 Then 'If second is 0
    b2 = "00" 'Assign second as 0 for display
End If
If b1 < 10 Then 'If second is less than 10
    b2 = "0" & b1 'Insert 0 before second for display
Else 'If second is not less than 10
    b2 = b1 'Assign second for display
End If
If a1 = 0 Then 'If minute is 0
    a2 = "00" 'Assign minute as 0 for display
End If
lblTimeLeft.Caption = a2 & ":" & b2 & ".000" 'Display timer

Exit Sub

ErrorHandler: 'Handles error
If b1 = 0 Then 'If second is 0
    txtSecond.Text = "" 'Display second as empty
Else 'If second is filled
    txtSecond.Text = b1 'Disable value change
End If
Resume Code: 'Run program code

End Sub
