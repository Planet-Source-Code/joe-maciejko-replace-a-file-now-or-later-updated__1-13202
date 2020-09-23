VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Replace v1.7"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "fmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   75
      TabIndex        =   17
      Top             =   3000
      Width           =   4515
      Begin VB.CheckBox chkExit 
         Caption         =   "Exit after process completes successfully"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   150
         Width           =   3840
      End
      Begin VB.TextBox txtMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1965
         TabIndex        =   5
         Text            =   "30"
         Top             =   420
         Width           =   315
      End
      Begin VB.CheckBox chkMinute 
         Caption         =   "Try again every       minute(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   6
         Top             =   450
         Width           =   3165
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1965
      Left            =   150
      TabIndex        =   10
      Top             =   0
      Width           =   4365
      Begin VB.TextBox txtExistingFile 
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   12
         Top             =   375
         Width           =   3615
      End
      Begin VB.TextBox txtReplaceFile 
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   975
         Width           =   3615
      End
      Begin VB.CommandButton cmdExistingFile 
         Caption         =   "..."
         Height          =   315
         Left            =   3825
         TabIndex        =   0
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton cmdReplaceFile 
         Caption         =   "..."
         Height          =   315
         Left            =   3825
         TabIndex        =   1
         Top             =   975
         Width           =   315
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   975
         TabIndex        =   2
         Top             =   1500
         Width           =   1065
      End
      Begin VB.TextBox txtTime 
         Height          =   315
         Left            =   3075
         TabIndex        =   3
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Existing File"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File to Rename"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label lblAltTimeDesc 
         AutoSize        =   -1  'True
         Caption         =   "Start Time:"
         Height          =   195
         Left            =   2250
         TabIndex        =   14
         Top             =   1575
         Width           =   765
      End
      Begin VB.Label lblAltDateDesc 
         AutoSize        =   -1  'True
         Caption         =   "Start Date:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1575
         Width           =   765
      End
   End
   Begin MSComDlg.CommonDialog CommDiag 
      Left            =   4050
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   2550
      TabIndex        =   8
      Top             =   2100
      Width           =   1065
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   390
      Left            =   1125
      TabIndex        =   7
      Top             =   2100
      Width           =   990
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   225
      Top             =   2025
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   375
      TabIndex        =   9
      Top             =   2700
      Width           =   3915
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ErrorFlag As Boolean
Dim lErrorNumber As Long
Dim bDateChanged As Boolean
Const sAppTitle = "File Replace"

Private Sub chkExit_Click()
    If chkExit.Value = vbChecked Then
        chkMinute.Value = vbChecked
    End If
End Sub

Private Sub cmdExistingFile_Click()

    CommDiag.InitDir = "C:\"
    CommDiag.ShowOpen
    
    txtExistingFile.Text = CommDiag.FileName
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
On Error GoTo ProcessErr
Dim Retval As Integer

    If cmdProcess.Caption = "Process" Then

        If txtExistingFile.Text = "" Then
            MsgBox "Select Existing File.", vbCritical, sAppTitle
            Exit Sub
        End If
    
        If txtReplaceFile.Text = "" Then
            MsgBox "Select Replacement File.", vbCritical, sAppTitle
            Exit Sub
        End If
    
        If txtDate.Text = "" Then
            MsgBox "Select Start Date.", vbCritical, sAppTitle
            Exit Sub
        End If
    
        If txtTime.Text = "" Then
            MsgBox "Select Start Time.", vbCritical, sAppTitle
            Exit Sub
        End If
                    
        lblStatus.Caption = ""
                        
        Retval = MsgBox("Start File Replace NOW?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title)
        If Retval = vbYes Then
            Timer1.Enabled = False
            ''Attempt to delete the existing file
            Kill txtExistingFile.Text
            ''If an error occured, display a message on the screen
            If ErrorFlag = True Then
                lblStatus.Caption = GetErrorName(lErrorNumber, True)
                ErrorFlag = False
            Else
                ''Attempt to rename the existing file
                Name txtReplaceFile.Text As txtExistingFile.Text
                ''If an error occured, display a message on the screen
                If ErrorFlag = True Then
                    lblStatus.Caption = GetErrorName(lErrorNumber, False)
                    ErrorFlag = False
                Else
                    ''If no errors then exit or display a message on the screen
                    If chkExit.Value = vbChecked Then
                        cmdExit_Click
                    Else
                        lblStatus.Caption = "File replaced!"
                    End If
                End If
                
            End If
        Else
            ''Set the timer
            Frame1.Enabled = False
            Frame2.Enabled = False
            Timer1.Enabled = True
            lblStatus.Caption = "Waiting For Start Time ..."
            cmdProcess.Caption = "Pause"
        End If
    ''Pause routine
    Else
        Frame1.Enabled = True
        Frame2.Enabled = True
        Timer1.Enabled = False
        lblStatus.Caption = ""
        cmdProcess.Caption = "Process"
    End If
    
    Exit Sub
    
ProcessErr:
    lErrorNumber = 0
    lErrorNumber = Err.Number
    ErrorFlag = True
    Resume Next
    
End Sub

Private Sub cmdReplaceFile_Click()

    CommDiag.InitDir = "C:\"
    CommDiag.ShowOpen
    
    txtReplaceFile.Text = CommDiag.FileName
    
End Sub

Private Sub Timer1_Timer()
On Error GoTo TimerErr

DoEvents

If DateDiff("h", Time, CDate(txtTime.Text)) <= 0 And _
   DateDiff("n", Time, CDate(txtTime.Text)) <= 0 And _
   DateDiff("d", Date, CDate(txtDate.Text)) <= 0 Then
                            
    'Only timer off
    Timer1.Enabled = False
    ''Attempt to delete the existing file
    Kill txtExistingFile.Text
    ''If an error occured, display a message on the screen
    If ErrorFlag = True Then
        ''Run if timer needs to be reset
        If chkMinute.Value = vbChecked Then
            ''If file is locked reset timer
            If lErrorNumber = 70 Or lErrorNumber = 75 Then
                Timer1.Enabled = True
                lblStatus.Caption = "Waiting For NEW Start Time ..."
                txtTime.Text = Format(DateAdd("n", CInt(txtMinute.Text), txtTime.Text), "hh:nn:ss AMPM")
            ''If file is not locked then display message to the screen
            Else
                Frame1.Enabled = True
                Frame2.Enabled = True
                Timer1.Enabled = False
                lblStatus.Caption = "Timer disabled.  " & GetErrorName(lErrorNumber, True)
                cmdProcess.Caption = "Process"
                chkMinute.Value = vbUnchecked
                chkExit.Value = vbUnchecked
            End If
        ''If an error occured, display a message on the screen
        Else
            Frame1.Enabled = True
            Frame2.Enabled = True
            Timer1.Enabled = False
            lblStatus.Caption = GetErrorName(lErrorNumber, True)
            cmdProcess.Caption = "Process"
        End If
        ErrorFlag = False
    Else
        ''Attempt to rename the existing file
        Name txtReplaceFile.Text As txtExistingFile.Text
        ''If an error occured, display a message on the screen
        If ErrorFlag = True Then
            Frame1.Enabled = True
            Frame2.Enabled = True
            Timer1.Enabled = False
            lblStatus.Caption = GetErrorName(lErrorNumber, False)
            cmdProcess.Caption = "Process"
            ErrorFlag = False
        ''If no errors then exit or display a message on the screen
        Else
            If chkExit.Value = vbChecked Then
                cmdExit_Click
            Else
                Frame1.Enabled = True
                Frame2.Enabled = True
                Timer1.Enabled = False
                lblStatus.Caption = "File replaced!"
                cmdProcess.Caption = "Process"
            End If
        End If
    End If
End If

Exit Sub

TimerErr:
    lErrorNumber = 0
    lErrorNumber = Err.Number
    ErrorFlag = True
    Resume Next
    
End Sub

Private Sub txtDate_LostFocus()
    txtDate.Text = Format(txtDate.Text, "mm/dd/yyyy")
End Sub

Private Sub txtMinute_GotFocus()
    txtMinute.SelStart = 0
    txtMinute.SelLength = Len(txtMinute.Text)
End Sub

Private Sub txtTime_Change()
    If Left(txtTime.Text, 2) = "12" And Right(txtTime.Text, 2) = "AM" Then
        If bDateChanged = False Then
            txtDate.Text = DateAdd("d", 1, txtDate.Text)
        End If
        bDateChanged = True
    Else
        bDateChanged = False
    End If
End Sub

Private Sub txtTime_LostFocus()
    txtTime.Text = Format(txtTime.Text, "hh:nn:ss AMPM")
End Sub

Private Function GetErrorName(lErrorNumber As Long, bExistingFileFlag As Boolean) As String
Dim sFileType As String

    If bExistingFileFlag = True Then
        sFileType = "Existing"
    Else
        sFileType = "Replacement"
    End If

    Select Case lErrorNumber
        Case 70
            GetErrorName = sFileType & " file is locked or open."
        Case 53
            GetErrorName = sFileType & " file not found."
        Case 75
            GetErrorName = sFileType & " file is locked or open."
        Case 58
            GetErrorName = sFileType & " file already exists."
        Case Else
            GetErrorName = Err.Number & " " & Err.Description
    End Select
    
End Function
