VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_main 
   Caption         =   "  Ion Beam Profiler v 0.1"
   ClientHeight    =   6150
   ClientLeft      =   2535
   ClientTop       =   2115
   ClientWidth     =   6180
   Icon            =   "x-y motor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   6180
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6480
      TabIndex        =   63
      Text            =   "&HFE"
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   8280
      TabIndex        =   62
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   6480
      TabIndex        =   61
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame11 
      Caption         =   "SAMPLE "
      Height          =   1215
      Left            =   4320
      TabIndex        =   57
      Top             =   1080
      Width           =   1335
      Begin VB.ComboBox cbo_sample_data_point 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         ItemData        =   "x-y motor.frx":0BC2
         Left            =   240
         List            =   "x-y motor.frx":0BE4
         TabIndex        =   58
         Text            =   "1"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "data/point"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame Frame10 
      Height          =   855
      Left            =   360
      TabIndex        =   52
      Top             =   5040
      Width           =   5295
      Begin VB.CommandButton cmd_run 
         Caption         =   "&RUN"
         Height          =   490
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   890
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&CANCEL"
         Enabled         =   0   'False
         Height          =   490
         Left            =   1320
         MousePointer    =   1  'Arrow
         TabIndex        =   55
         Top             =   240
         Width           =   890
      End
      Begin VB.CommandButton cmd_setup 
         Caption         =   "SETUP >>"
         Height          =   490
         Left            =   4080
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "E&XIT"
         Height          =   490
         Left            =   2520
         TabIndex        =   53
         Top             =   240
         Width           =   890
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2415
      Left            =   2760
      TabIndex        =   36
      Top             =   2520
      Width           =   2895
      Begin VB.Frame Frame9 
         Caption         =   "z - axis"
         Height          =   975
         Left            =   0
         TabIndex        =   49
         Top             =   1440
         Width           =   2895
         Begin VB.CommandButton cmd_perset_z_axis 
            Caption         =   "PRESET"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   51
            Top             =   380
            Width           =   735
         End
         Begin VB.ComboBox cbo_z_post 
            Height          =   330
            ItemData        =   "x-y motor.frx":0C07
            Left            =   240
            List            =   "x-y motor.frx":0C29
            TabIndex        =   50
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "mm"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   60
            Top             =   380
            Width           =   255
         End
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   550
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   38
         Top             =   1080
         Width           =   550
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2000
         TabIndex        =   37
         Top             =   1080
         Width           =   550
      End
      Begin VB.Label lbl_x_post 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   100
         TabIndex        =   48
         Top             =   600
         Width           =   550
      End
      Begin VB.Label lbl_y_post 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1050
         TabIndex        =   47
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lbl_z_post 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2000
         TabIndex        =   46
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "x-position"
         Height          =   300
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "y-position"
         Height          =   300
         Left            =   1080
         TabIndex        =   44
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "z-position"
         Height          =   300
         Left            =   2040
         TabIndex        =   43
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Left            =   705
         TabIndex        =   42
         Top             =   620
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Left            =   2600
         TabIndex        =   41
         Top             =   620
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Left            =   1650
         TabIndex        =   40
         Top             =   620
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cmdlog1 
      Left            =   0
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "SETUP"
      Height          =   3495
      Left            =   6240
      TabIndex        =   26
      Top             =   120
      Width           =   2415
      Begin VB.Frame Frame6 
         Caption         =   "INSTUMENT ADDR"
         Height          =   855
         Left            =   240
         TabIndex        =   33
         Top             =   2400
         Width           =   1935
         Begin VB.ComboBox cbo_dmm_addr 
            Height          =   330
            ItemData        =   "x-y motor.frx":0C54
            Left            =   840
            List            =   "x-y motor.frx":0CB5
            TabIndex        =   34
            Text            =   "22"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "DMM"
            Height          =   255
            Left            =   195
            TabIndex        =   35
            Top             =   400
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "HPIB ADDR"
         Height          =   855
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   1935
         Begin VB.ComboBox cbo_hpib_addr 
            Height          =   330
            ItemData        =   "x-y motor.frx":0D2B
            Left            =   840
            List            =   "x-y motor.frx":0D5C
            TabIndex        =   31
            Text            =   "7"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "HPIB"
            Height          =   255
            Left            =   195
            TabIndex        =   32
            Top             =   400
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "IO PORT ADDR"
         Height          =   855
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1935
         Begin VB.ComboBox cbo_ioport_addr 
            Height          =   330
            ItemData        =   "x-y motor.frx":0D92
            Left            =   840
            List            =   "x-y motor.frx":0D9C
            TabIndex        =   28
            Text            =   "LPT1"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "LPT"
            Height          =   255
            Left            =   195
            TabIndex        =   29
            Top             =   400
            Width           =   810
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000000&
      Caption         =   "Set Positions of Motor"
      Height          =   2415
      Left            =   360
      TabIndex        =   19
      Top             =   2520
      Width           =   5295
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   495
         Index           =   5
         Left            =   150
         Picture         =   "x-y motor.frx":0DAC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "move forward"
         Top             =   1800
         Width           =   900
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   495
         Index           =   4
         Left            =   1400
         Picture         =   "x-y motor.frx":11EE
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "move backward"
         Top             =   1800
         Width           =   900
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   630
         Index           =   2
         Left            =   840
         Picture         =   "x-y motor.frx":1630
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "move up"
         Top             =   360
         Width           =   750
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   1335
         Index           =   0
         Left            =   150
         Picture         =   "x-y motor.frx":1A72
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "move left"
         Top             =   360
         Width           =   600
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   1335
         Index           =   1
         Left            =   1680
         Picture         =   "x-y motor.frx":1EB4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "move right"
         Top             =   360
         Width           =   600
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   630
         Index           =   3
         Left            =   840
         Picture         =   "x-y motor.frx":22F6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "move down"
         Top             =   1080
         Width           =   750
      End
   End
   Begin VB.PictureBox pct_area 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   18
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCAN AERA"
      Height          =   1215
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   3855
      Begin VB.ComboBox cbo_x_min 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor.frx":2738
         Left            =   960
         List            =   "x-y motor.frx":274B
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbo_x_max 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor.frx":2761
         Left            =   2280
         List            =   "x-y motor.frx":2774
         TabIndex        =   10
         Text            =   "20"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbo_y_min 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor.frx":278A
         Left            =   960
         List            =   "x-y motor.frx":279D
         TabIndex        =   9
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cbo_y_max 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor.frx":27B3
         Left            =   2280
         List            =   "x-y motor.frx":27C6
         TabIndex        =   8
         Text            =   "20"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "x axis"
         Height          =   210
         Index           =   8
         Left            =   375
         TabIndex        =   17
         Top             =   320
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Index           =   9
         Left            =   3375
         TabIndex        =   16
         Top             =   320
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "y axis"
         Height          =   210
         Index           =   10
         Left            =   375
         TabIndex        =   15
         Top             =   800
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Index           =   11
         Left            =   3375
         TabIndex        =   14
         Top             =   800
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "to"
         Height          =   210
         Index           =   12
         Left            =   1920
         TabIndex        =   13
         Top             =   320
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "to"
         Height          =   210
         Index           =   13
         Left            =   1920
         TabIndex        =   12
         Top             =   800
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCAN RESULATION"
      Height          =   855
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox cbo_x_step 
         BackColor       =   &H00C0C0FF&
         Height          =   330
         ItemData        =   "x-y motor.frx":27DC
         Left            =   960
         List            =   "x-y motor.frx":27FB
         TabIndex        =   2
         Text            =   "0.10"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cbo_y_step 
         BackColor       =   &H00C0C0FF&
         Height          =   330
         ItemData        =   "x-y motor.frx":2835
         Left            =   3600
         List            =   "x-y motor.frx":2854
         TabIndex        =   1
         Text            =   "0.10"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Index           =   2
         Left            =   1950
         TabIndex        =   6
         Top             =   400
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Index           =   3
         Left            =   4650
         TabIndex        =   5
         Top             =   400
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "step x"
         Height          =   210
         Index           =   4
         Left            =   150
         TabIndex        =   4
         Top             =   400
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "step y"
         Height          =   210
         Index           =   5
         Left            =   2880
         TabIndex        =   3
         Top             =   400
         Width           =   435
      End
   End
   Begin MSCommLib.MSComm Serial 
      Left            =   0
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private flag_plot As Boolean

Private data_port As Integer
Private control_port As Integer
Private hpib_dmm As String
Private data_sample_count As String

Private Const x_axis = 0   'Pin 1
Private Const y_axis = 1   'Pin 14
Private Const z_axis = 2   'Pin 16

Private x_post As Single
Private y_post As Single
Private z_post As Single

Private step_fwd(0 To 9) As Byte
Private step_rwd(0 To 9) As Byte
Private m_speed As Integer
Private m_step As Integer
'unsigned int mot_step, mot_speed;
'unsigned int x_step,y_step;
'unsigned int x_min,x_max;
'unsigned int y_min,y_max;
Private id As Integer     ' device session id

Private Const x_left = 1

Private x_cord1 As Integer
Private y_cord1 As Integer
Private x_cord2 As Integer
Private y_cord2 As Integer
'Private DataP As Integer
Private motor_flag As Boolean

Private Sub save_data_file(data_read_xy() As Single, x_max As Integer, y_max As Integer)
    Dim save_cancel As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim data_buff As String
    
    On Error GoTo ErrHandler
    
    save_cancel = False
    
    With cmdlog1
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt Or cdlOFNReadOnly 'cdlOFNNoChangeDir
        .Filter = "Text Files(*.txt)|*.txt|Ion Source Files(*.xls)|*.xls"
        .FilterIndex = 2
        .ShowSave
    End With
    
    If Not save_cancel Then
        Open cmdlog1.FileName For Output As #1
        For j = 0 To y_max
          data_buff = Format(data_read_xy()(j, 0), "0.00000000")
          For i = 1 To x_max
               data_buff = data_buff & "," & Format(data_read_xy()(j, i), "0.00000000")
          Next i
          Print #1, data_buff
        Next j
        Close #1
    End If
    
ErrHandler:
  If Err.Number = cdlCancel Then
    save_cancel = True
    Resume Next
  End If
End Sub
Private Function str2sng(data_read_xy_s As String) As Single
    Dim i As Integer
    Dim data_read_array() As String
    Dim data_read As Single
    
    data_read_array = Split(data_read_xy_s, ",")
    data_read = 0
    For i = LBound(data_read_array) To UBound(data_read_array)
        data_read = data_read + Val(data_read_array(i))
    Next i
    str2sng = data_read / (UBound(data_read_array) - LBound(data_read_array) + 1)
End Function

Private Function get_ltp_addr(ltp_port As String) As Integer
  Select Case ltp_port
    Case "LPT1": get_ltp_addr = &H378
    Case "LPT2": get_ltp_addr = &H278
  End Select
End Function
Private Sub get_setup_parameter()
  data_port = get_ltp_addr(cbo_ioport_addr.Text)
  control_port = data_port + 2
    
  'If id <> 0 Then
   '   iclose (id)
  'End If
  hpib_dmm = "hpib" & cbo_hpib_addr.Text & "," & cbo_dmm_addr.Text
  
  data_sample_count = "SAMP:COUN " & cbo_sample_data_point.Text + Chr$(10)
End Sub

Private Sub update_xyz_post(ByRef axis_post As Single, axis_step As Single, axis_post_disp As Label)
  
  axis_post = axis_post + axis_step
  axis_post_disp.Visible = False
  axis_post_disp.Caption = Format(axis_post, "00.00")
  axis_post_disp.Visible = True
End Sub

Private Function read_hpib() As String
                  
   Dim readbuf As String * 200      ' buffer used for iread
   Dim nargs As Integer             ' # args converted by format string

'  Set up an error handler within this subroutine that will get
'  called if a SICL error occurs.
   On Error GoTo ErrorHandler

   'nargs = ivprintf(id, "READ?" + Chr$(10))
   nargs = ivprintf(id, "INIT" + Chr$(10))
   nargs = ivprintf(id, "*TRG" + Chr$(10))
   nargs = ivprintf(id, "FETC?" + Chr$(10))
   
   nargs = ivscanf(id, "%200t", readbuf)

'  Close the device session.
 '  Call iclose(id)
   read_hpib = readbuf
   Exit Function

ErrorHandler:
'  Close the device session if iopen was successful.
   If id <> 0 Then
      iclose (id)
   End If
   MsgBox Error$, vbOKOnly, "ERROR"
   read_hpib = "0"
   Exit Function
End Function

Private Function initialize_hpib() As Boolean
                  
   'Dim readbuf As String * 200      ' buffer used for iread
   'Dim commandstr As String * 128   ' command passed to instrument
   'Dim index As Integer             ' used to parse SCPI error message
   Dim nargs As Integer             ' # args converted by format string

'  Set up an error handler within this subroutine that will get
'  called if a SICL error occurs.
   Call get_setup_parameter
   
   On Error GoTo ErrorHandler

'  Open a device session using the device address contained in
   'If id = 0 Then
    id = iopen(hpib_dmm)
   'End If
   'Me.Caption = id
'  Set the I/O timeout value for this session to 1 second.
    Call itimeout(id, 10000)
    
'  Clear the error/event queue for the instrument.  This allows
'  us to query the instrument after sending a command to see if
'  the command was accepted.
   nargs = ivprintf(id, "*CLS" & vbLf)
   
   nargs = ivprintf(id, "CONF:volt:DC 10,0.0001" + Chr$(10)) ' uncomment this line for voltage
   'nargs = ivprintf(id, "volt:DC:nplc 10" + Chr$(10))
   
   'nargs = ivprintf(id, "CONF:CURR:DC 0.0001,0.0000001" & vbLf)    ' uncomment this line for current
   'nargs = ivprintf(id, "CURR:DC:nplc 10" & vbLf)                  ' uncomment this line for current
   
   nargs = ivprintf(id, "TRIG:SOUR BUS" & vbLf)
   nargs = ivprintf(id, data_sample_count)
      
   nargs = ivprintf(id, "TRIG:DELAY 2" & vbLf)                       'delay 2 s for volt   ---- 0 for current
      
   'nargs = ivprintf(id, "TRIG:COUN 1" + Chr$(10))
   'nargs = ivprintf(id, "*CLS" + Chr$(10))
   

'  Close the device session.
 '  Call iclose(id)
   initialize_hpib = True
   Exit Function

ErrorHandler:
'  Close the device session if iopen was successful.
   If id <> 0 Then
      iclose (id)
   End If
   MsgBox Error$, vbOKOnly, "SOMETHING WRONG"
   initialize_hpib = False
   Exit Function
End Function

Private Sub scan(ByRef data_read_xy() As Single, x_step As Integer, x_max As Integer, _
                 y_step As Integer, y_max As Integer)
  
  Dim i As Integer
  Dim j As Integer
  Dim data_read_xy_s() As String * 200
  'Dim data_read_xy() As Single
  Dim data_x As String
  
  Dim x_post_step As Single
  Dim y_post_step As Single
  ReDim data_read_xy_s(0 To y_max, 0 To x_max)
  'ReDim data_read_xy(0 To y_max, 0 To x_max)
  
  x_post_step = CSng(cbo_x_step.Text)        'convert distance to step(1 mm = 500 step)
  y_post_step = CSng(cbo_y_step.Text)
  
  flag_plot = True
  
  Me.MousePointer = 11
  
  For j = 0 To y_max
   
   For i = 0 To x_max - 1
      If (j Mod 2) = 0 Then
        
        data_read_xy_s(j, i) = read_hpib()
        
        Call xyz_move(step_rwd(), x_axis, x_step, m_speed)
        Call update_xyz_post(x_post, x_post_step, lbl_x_post)
      Else
        data_read_xy_s(j, x_max - i) = read_hpib()
        Call xyz_move(step_fwd(), x_axis, x_step, m_speed)
        Call update_xyz_post(x_post, -x_post_step, lbl_x_post)
      End If
      
      If Not flag_plot Then GoTo abort_scan
      vbrcDelayTime 1
   Next i
   
   If (j Mod 2) = 0 Then
     data_read_xy_s(j, x_max) = read_hpib()
   Else
     data_read_xy_s(j, 0) = read_hpib()
   End If
     
   If j <> y_max Then
     Call motor_off
     Call xyz_move(step_rwd(), y_axis, y_step, m_speed)
     Call update_xyz_post(y_post, y_post_step, lbl_y_post)
     Call motor_off
   End If
  Next j
  'save file
  'Me.Caption = "scan complete"
abort_scan: 'CALL WRITE FILE
  
  For j = 0 To y_max
    For i = 0 To x_max
      data_read_xy(j, i) = str2sng(data_read_xy_s(j, i))
    Next i
  Next j
  
  Me.MousePointer = 0
End Sub

Private Sub preset_scan(x_min As Integer, y_min As Integer)
  Dim x_step As Integer
  Dim y_step As Integer
  
  x_step = x_min * 500 '***x_min-x_post****
  y_step = y_min * 500
  
  Call xyz_move(step_rwd(), x_axis, x_step, m_speed)
  Call motor_off
  x_post = x_min
  Call update_xyz_post(x_post, 0, lbl_x_post)
 Call update_xyz_post(y_post, 0, lbl_y_post)
  Call xyz_move(step_rwd(), y_axis, y_step, m_speed)
  Call motor_off
  y_post = y_min
  Call update_xyz_post(y_post, 0, lbl_y_post)
End Sub

Private Sub scan_data()
 Dim x_step As Integer
 Dim y_step As Integer
 Dim x_min As Integer
 Dim y_min As Integer
 Dim x_max As Integer
 Dim y_max As Integer
 Dim data_read_xy() As Single
 'Dim m_x_step As Integer
 'Dim m_y_step As Integer
 
 x_step = Val(cbo_x_step.Text) * 500        'convert distance to step(1 mm = 500 step)
 y_step = Val(cbo_y_step.Text) * 500
 x_min = Val(cbo_x_min.Text)
 x_max = Val(cbo_x_max.Text)
 y_min = Val(cbo_y_min.Text)
 y_max = Val(cbo_y_max.Text)
 
 x_max = (x_max - x_min) * 500 / x_step
 y_max = (y_max - y_min) * 500 / y_step
 

 Call preset_scan(x_min, y_min)
 ReDim data_read_xy(0 To y_max, 0 To x_max)
 Call scan(data_read_xy(), x_step, x_max, y_step, y_max)
 Call save_data_file(data_read_xy(), x_max, y_max)
End Sub
Private Sub xyz_move(motor_phase() As Byte, motor_axis As Byte, _
            motor_step As Integer, motor_speed)
  
  Dim i As Integer
  
  For i = 0 To motor_step - 1
      Out data_port, motor_phase(i Mod 10)
      NotPortBit control_port, motor_axis
      NotPortBit control_port, motor_axis
      NotPortBit control_port, motor_axis
      'Out control_port, &H0
      'Out control_port, motor_axis
      'Out control_port, &H0
      vbrcDelayTime motor_speed
      DoEvents
  Next
End Sub
Private Sub initialize_motor()
 
 step_fwd(0) = &H12
 step_fwd(1) = &H16
 step_fwd(2) = &H14
 step_fwd(3) = &H15
 step_fwd(4) = &H5
 step_fwd(5) = &HD
 step_fwd(6) = &H9
 step_fwd(7) = &HB
 step_fwd(8) = &HA
 step_fwd(9) = &H1A
 
 step_rwd(0) = &H1A
 step_rwd(1) = &HA
 step_rwd(2) = &HB
 step_rwd(3) = &H9
 step_rwd(4) = &HD
 step_rwd(5) = &H5
 step_rwd(6) = &H15
 step_rwd(7) = &H14
 step_rwd(8) = &H16
 step_rwd(9) = &H12
 
 m_speed = 1
 m_step = 10
 Call get_setup_parameter
 Call motor_off
 
 x_post = 0
 y_post = 0
 z_post = 0
  
 Call update_xyz_post(x_post, 0, lbl_x_post)
 Call update_xyz_post(y_post, 0, lbl_y_post)
 Call update_xyz_post(z_post, 0, lbl_z_post)
End Sub
Private Sub motor_off()
  Out data_port, &HFF
  ClrPortBit control_port, x_axis
  ClrPortBit control_port, y_axis
  SetPortBit control_port, z_axis
  'vbrcDelayTime 500
  SetPortBit control_port, x_axis
  SetPortBit control_port, y_axis
  ClrPortBit control_port, z_axis
  'vbrcDelayTime 500
  ClrPortBit control_port, x_axis
  ClrPortBit control_port, y_axis
  SetPortBit control_port, z_axis
  
End Sub

'Private btn_flag As Boolean

Private Sub motor_move(axis As Integer)
    Dim m_phase() As Byte
    Dim m_axis As Byte
    
    Dim axis_post As Single
    Dim axis_step As Single
    Dim axis_post_disp As Label
    
    Call motor_off
    motor_flag = True
    
    Select Case axis
        Case 0: m_axis = x_axis: m_phase() = step_fwd(): _
                Set axis_post_disp = lbl_x_post: _
                axis_post = x_post: axis_step = -0.02 'move left
        Case 1: m_axis = x_axis: m_phase() = step_rwd(): _
                Set axis_post_disp = lbl_x_post: _
                axis_post = x_post: axis_step = 0.02  'move right
        Case 2: m_axis = y_axis: m_phase() = step_fwd(): _
                Set axis_post_disp = lbl_y_post: _
                axis_post = y_post: axis_step = -0.02 'move up
        Case 3: m_axis = y_axis: m_phase() = step_rwd(): _
                Set axis_post_disp = lbl_y_post: _
                axis_post = y_post: axis_step = 0.02  'move down
        Case 4: m_axis = z_axis: m_phase() = step_fwd(): _
                Set axis_post_disp = lbl_z_post: _
                axis_post = z_post: axis_step = 0.02  'move back
        Case 5: m_axis = z_axis: m_phase() = step_rwd(): _
                Set axis_post_disp = lbl_z_post: _
                axis_post = z_post: axis_step = -0.02  'move front
    End Select
    
    Do While motor_flag
        Call update_xyz_post(axis_post, axis_step, axis_post_disp)
        Call xyz_move(m_phase(), m_axis, 10, m_speed)
        DoEvents
    Loop
    
    Select Case axis
        Case 0, 1: x_post = axis_post
        Case 2, 3: y_post = axis_post
        Case 4, 5: z_post = axis_post
    End Select
    Call motor_off
End Sub


Private Sub cmd_exit_Click()
  Unload Me
End Sub

Private Sub cmd_mot_mov_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   motor_move (Index)
End Sub

Private Sub cmd_mot_mov_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    motor_flag = False
End Sub



Private Sub cmd_perset_z_axis_Click()
    Dim new_z_post As Single
    Dim i As Integer
    Dim j As Integer
    Dim m_phase() As Byte
    Dim axis_step As Single
    
    new_z_post = CSng(cbo_z_post.Text)
    If new_z_post > z_post Then
        m_phase() = step_fwd()
        axis_step = 0.02
    Else
        m_phase() = step_rwd()
        axis_step = -0.02
    End If
    
    j = Abs(new_z_post - z_post) / 0.02
    
    For i = 1 To j
     Call update_xyz_post(z_post, axis_step, lbl_z_post)
     Call xyz_move(m_phase(), z_axis, 10, m_speed)
    Next i
    Call motor_off
    
End Sub

Private Sub cmd_reset_Click(Index As Integer)
  Dim idx As Integer
   
  idx = Index
  Select Case idx
   Case 0: x_post = 0
   Case 1: y_post = 0
   Case 2: z_post = 0
   Case 3: x_post = 0: y_post = 0: z_post = 0
  End Select
  
  Call update_xyz_post(x_post, 0, lbl_x_post)
  Call update_xyz_post(y_post, 0, lbl_y_post)
  Call update_xyz_post(z_post, 0, lbl_z_post)
End Sub

Private Sub cmd_run_Click()
  If Not initialize_hpib Then
   'MsgBox "NOT OK"
   Exit Sub
  End If
  'MsgBox "Scan Now"
  cmd_run.Enabled = False
  cmd_exit.Enabled = False
  cmd_cancel.Enabled = True
  Call scan_data
  cmd_run.Enabled = True
  cmd_exit.Enabled = True
  cmd_cancel.Enabled = False
  Call motor_off
  'MsgBox "Scan Complete"
End Sub

Private Sub cmd_cancel_Click()
  If Not cmd_run.Enabled Then
   If MsgBox("Abort Scan?", vbOKCancel, "Scanning pause") = vbOK Then
    Me.MousePointer = 1
    flag_plot = False
    cmd_run.Enabled = True
    cmd_cancel.Enabled = False
   End If
  End If
End Sub

Private Sub cmd_setup_Click()
  If cmd_setup.Caption = "SETUP >>" Then
    cmd_setup.Caption = "SETUP <<"
    Me.width = 9000
  Else
    cmd_setup.Caption = "SETUP >>"
    Me.width = 6200
  End If
End Sub



Private Sub Command1_Click()
    Out data_port, Val(Text1.Text)
      Out control_port, &H0
      Out control_port, z_axis
      Out control_port, &H0
End Sub

Private Sub Command2_Click()
Call motor_off
End Sub



Private Sub Form_Load()
  Call initialize_motor
  'Call initialize_hpib
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("Exit Program ?", vbOKCancel) = vbCancel Then
    Cancel = 1
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbNormal Then
   Me.Move (Screen.width - Me.width) / 2, (Screen.Height - Me.Height) / 2
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call initialize_motor
   Call siclcleanup    ' Tell SICL to clean up for this task
   End
End Sub

Private Sub pct_area_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If Not pct_area.AutoRedraw Then
            pct_area.AutoRedraw = True
            pct_area.Cls
            pct_area.AutoRedraw = False
        End If
        pct_area.Cls
        x_cord1 = X
        y_cord1 = Y
    End If
End Sub

Private Sub pct_area_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = "x= " & X & "   y= " & Y
    
    If Button = vbLeftButton Then
      pct_area.Refresh
      x_cord2 = X
      y_cord2 = Y
      
      pct_area.Line (x_cord1, x_cord1)-(x_cord2, x_cord2), , B
    
    Else
     
     pct_area.Refresh
     pct_area.Line (0, Y)-(pct_area.ScaleWidth, Y)
     pct_area.Line (X, 0)-(X, pct_area.ScaleHeight)

    End If
End Sub

Private Sub pct_area_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
     pct_area.Refresh
     pct_area.AutoRedraw = True
     x_cord2 = X
     y_cord2 = Y
     
     pct_area.Line (x_cord1, x_cord1)-(x_cord2, x_cord2), , B
     pct_area.AutoRedraw = False
    End If
End Sub
