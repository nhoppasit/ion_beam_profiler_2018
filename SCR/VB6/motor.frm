VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form displayScanData 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Data"
   ClientHeight    =   6900
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   11685
   Icon            =   "motor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   11685
   Begin VB.Frame Scan1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scanning Data"
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   855
         Left            =   4680
         TabIndex        =   62
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Scan Now"
         Height          =   735
         Left            =   3480
         TabIndex        =   45
         Top             =   5520
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   855
         Left            =   3480
         TabIndex        =   44
         Top             =   4440
         Width           =   975
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000000&
         Caption         =   "Set Positions of Motor"
         Height          =   1935
         Left            =   360
         TabIndex        =   39
         Top             =   4320
         Width           =   2895
         Begin VB.CommandButton Command16 
            BackColor       =   &H00FF80FF&
            Height          =   615
            Left            =   960
            Picture         =   "motor.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00FF80FF&
            Height          =   1335
            Left            =   2040
            Picture         =   "motor.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00FF80FF&
            Height          =   1335
            Left            =   240
            Picture         =   "motor.frx":09CE
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00FF80FF&
            Height          =   615
            Left            =   960
            Picture         =   "motor.frx":0E10
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Locate of XY Axis"
         Height          =   1095
         Index           =   0
         Left            =   5880
         TabIndex        =   36
         Top             =   360
         Width           =   5175
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Y-axis = 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Index           =   0
            Left            =   2880
            TabIndex        =   38
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X-axis = 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Index           =   1
            Left            =   840
            TabIndex        =   37
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scan Area of  XY Axis"
         Height          =   1215
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   2880
         Width           =   5295
         Begin VB.ComboBox Combo1 
            Height          =   330
            Index           =   6
            ItemData        =   "motor.frx":1252
            Left            =   3000
            List            =   "motor.frx":1262
            TabIndex        =   16
            Text            =   "20"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   330
            Index           =   5
            ItemData        =   "motor.frx":1275
            Left            =   960
            List            =   "motor.frx":1285
            TabIndex        =   15
            Text            =   "20"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   330
            Index           =   4
            ItemData        =   "motor.frx":1298
            Left            =   3000
            List            =   "motor.frx":12A8
            TabIndex        =   14
            Text            =   "5"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   330
            Index           =   3
            ItemData        =   "motor.frx":12BB
            Left            =   960
            List            =   "motor.frx":12CE
            TabIndex        =   13
            Text            =   "5"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "to"
            Height          =   210
            Index           =   13
            Left            =   2640
            TabIndex        =   22
            Top             =   840
            Width           =   165
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "to"
            Height          =   210
            Index           =   12
            Left            =   2640
            TabIndex        =   21
            Top             =   360
            Width           =   165
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "mm."
            Height          =   210
            Index           =   11
            Left            =   4545
            TabIndex        =   20
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Y axis"
            Height          =   210
            Index           =   10
            Left            =   360
            TabIndex        =   19
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "mm."
            Height          =   210
            Index           =   9
            Left            =   4545
            TabIndex        =   18
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "X axis"
            Height          =   210
            Index           =   8
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scale Scan of  XY Axis"
         Height          =   735
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   5295
         Begin VB.ComboBox Combo1 
            Height          =   330
            Index           =   2
            ItemData        =   "motor.frx":12E4
            Left            =   3240
            List            =   "motor.frx":12F7
            TabIndex        =   7
            Text            =   "0.10"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   330
            Index           =   1
            ItemData        =   "motor.frx":1319
            Left            =   960
            List            =   "motor.frx":132C
            TabIndex        =   6
            Text            =   "0.10"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Step Y"
            Height          =   210
            Index           =   5
            Left            =   2640
            TabIndex        =   11
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Step X"
            Height          =   210
            Index           =   4
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "mm."
            Height          =   210
            Index           =   3
            Left            =   4560
            TabIndex        =   9
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "mm."
            Height          =   210
            Index           =   2
            Left            =   2280
            TabIndex        =   8
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Setup Computer Port"
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   5295
         Begin VB.OptionButton Option2 
            Caption         =   "COM 4"
            Height          =   375
            Left            =   4080
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "COM 3"
            Height          =   375
            Left            =   2880
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM 2"
            Height          =   375
            Left            =   1680
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "COM 1"
            Height          =   375
            Left            =   480
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin VB.Frame Frame11 
         Caption         =   "Picture of Scan Equipotential System"
         Height          =   4575
         Left            =   5880
         TabIndex        =   24
         Top             =   1680
         Width           =   5175
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   360
            MousePointer    =   2  'Cross
            ScaleHeight     =   9.299
            ScaleMode       =   0  'User
            ScaleWidth      =   9.242
            TabIndex        =   25
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Y"
            Height          =   210
            Index           =   6
            Left            =   4920
            TabIndex        =   33
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "X"
            Height          =   210
            Index           =   7
            Left            =   2520
            TabIndex        =   32
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "X"
            Height          =   210
            Index           =   34
            Left            =   2520
            TabIndex        =   31
            Top             =   4200
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Y"
            Height          =   210
            Index           =   32
            Left            =   120
            TabIndex        =   30
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "(9,9)"
            Height          =   210
            Index           =   30
            Left            =   4635
            TabIndex        =   29
            Top             =   4200
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "(0,9)"
            Height          =   210
            Index           =   29
            Left            =   195
            TabIndex        =   28
            Top             =   4200
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "(9,0)"
            Height          =   210
            Index           =   28
            Left            =   4635
            TabIndex        =   27
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "(0,0)"
            Height          =   210
            Index           =   27
            Left            =   195
            TabIndex        =   26
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   1800
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6000
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7320
      Top             =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7680
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Scan3 
      BackColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   3240
      TabIndex        =   59
      Top             =   960
      Width           =   5055
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   420
         Index           =   1
         Left            =   1200
         TabIndex        =   61
         Text            =   "Motor Moving to Position"
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   420
         Index           =   2
         Left            =   1200
         TabIndex        =   60
         Top             =   960
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   735
         Left            =   240
         Shape           =   3  'Circle
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Scan2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data Input of Step Motor Scan from Port COM"
      Height          =   5535
      Left            =   2760
      TabIndex        =   46
      Top             =   600
      Width           =   6015
      Begin VB.CommandButton Command18 
         Caption         =   "< Pause >"
         Height          =   495
         Left            =   3120
         TabIndex        =   48
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton Command19 
         Caption         =   "< Continue >"
         Height          =   495
         Left            =   3120
         TabIndex        =   49
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "< Stop >"
         Height          =   495
         Left            =   1320
         TabIndex        =   47
         Top             =   4560
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   480
         TabIndex        =   50
         Top             =   2520
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1085
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   1320
         TabIndex        =   58
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Voltage (V)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   0
         Left            =   1320
         TabIndex        =   57
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remaining Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3240
         TabIndex        =   56
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   55
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Location of Scan (cm)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   720
         TabIndex        =   54
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   53
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y , X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   52
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   51
         Top             =   2640
         Width           =   855
      End
   End
End
Attribute VB_Name = "displayScanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyD, KeyU, KeyL, KeyR As Boolean
Dim ComPortValue As Integer
Dim coppic As Boolean
Dim x00 As Integer
Dim y00 As Integer
Dim x0 As Integer
Dim y0 As Integer
Dim sx1, sy1, sy2, sx2 As Integer
Dim DataP1000, DataP100, DataP10, DataP1 As Integer
Dim DataP As Currency
Dim CountXY As Integer
Dim HHH As Boolean
Dim CCC As Boolean
Dim kxy As Long
Dim xy As Long
Dim lll As Long
Dim xxx1 As Long
Dim hh, mm, ss, tt2, hh1 As Integer
Dim xxx, yyy As Currency
Dim hun As Integer
Dim ax1 As Long

Private Sub Combo1_Change(Index As Integer)
On Error Resume Next
    CheckInput
End Sub

Private Sub Combo1_Click(Index As Integer)
On Error Resume Next
    CheckInput
End Sub

Private Sub Command1_Click()
On Error Resume Next
    displayScanData.Hide
    Scan1.Visible = True
    Scan3.Visible = True
    Command3.Value = True
    For i = 2 To 3 Step 1
        mdiProjectEPL.mdiToolbar.Buttons(i).Enabled = True
    Next i
    mdiProjectEPL.menuOpen.Enabled = True
    mdiProjectEPL.mdiStatusBar.Panels(1) = "Status : " & "Done"
    Scan1.ZOrder
    ClearScan
End Sub

Private Sub Command18_Click()
On Error Resume Next
    Timer2.Enabled = False
    Command19.ZOrder
    HHH = True
    HHH1
End Sub

Private Sub Command19_Click()
On Error Resume Next
    Timer2.Enabled = True
    Command18.ZOrder
    HHH = False
    CCC = True
    CCC1
End Sub

Private Sub Command2_Click()
On Error Resume Next
      
        SetComPort
        DataToPort
      
End Sub

Private Sub Command3_Click()
On Error Resume Next
    Picture2.AutoRedraw = True
    Picture2.Cls
        For i = -1 To 8 Step 1
            Picture2.Line (i + 1, 0)-(9, 9), vbWhite, B
            Picture2.Line (0, i + 1)-(9, 9), vbWhite, B
            DoEvents
        Next i
    Picture2.AutoRedraw = False
    For i = 1 To 6 Step 1
        Combo1(i) = ""
    Next i
    Text26 = ""
    Text27 = ""
    Text28 = ""
    Text29 = ""
    Text30 = ""
    Text31 = ""
    Text32 = ""
    kxy = 0
    Option3.Value = True
End Sub

Private Sub Command6_Click()
On Error Resume Next
    MSComm1.Output = "x"
    Timer2.Enabled = False
    displayScanData.Hide
    Scan1.Visible = True
    Scan3.Visible = True
    Command3.Value = True
    For i = 2 To 3 Step 1
        mdiProjectEPL.mdiToolbar.Buttons(i).Enabled = True
    Next i
    mdiProjectEPL.menuOpen.Enabled = True
    mdiProjectEPL.mdiStatusBar.Panels(1) = "Status : " & "Done"
    If Timer1.Enabled Then
        Timer1.Enabled = False
    End If
    Scan1.ZOrder
    SaveData
End Sub

Private Sub Form_Activate()
On Error Resume Next
    mdiProjectEPL.mdiStatusBar.Panels(1) = "Status : " & "New Scan"
    For i = 2 To 17 Step 1
        mdiProjectEPL.mdiToolbar.Buttons(i).Enabled = False
    Next i
    mdiProjectEPL.menuOpen.Enabled = False
        Picture2.AutoRedraw = True
            For i = -1 To 8 Step 1
                Picture2.Line (i + 1, 0)-(9, 9), vbWhite, B
                Picture2.Line (0, i + 1)-(9, 9), vbWhite, B
            DoEvents
            Next i
        Picture2.AutoRedraw = False
    Nsquar = 0
    Ncircle = 0
    Option3.Value = True
    ClearScan
End Sub

Private Sub Form_Deactivate()
On Error Resume Next
    For i = 2 To 3 Step 1
        mdiProjectEPL.mdiToolbar.Buttons(i).Enabled = True
    Next i
    mdiProjectEPL.menuOpen.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.WindowState = 2
    Timer1.Enabled = False
    Command2.Enabled = False
    kxy = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    For i = 2 To 3 Step 1
        mdiProjectEPL.mdiToolbar.Buttons(i).Enabled = True
    Next i
    mdiProjectEPL.menuOpen.Enabled = True
    mdiProjectEPL.mdiStatusBar.Panels(1) = "Status : " & "Done"
    Exit Sub
    ClearScan
End Sub

Private Sub Command13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Not MSComm1.PortOpen Then
        MSComm1.PortOpen = True
    End If
    KeyU = True
    KeyUp
End Sub

Private Sub Command13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    KeyU = False
    MSComm1.PortOpen = False
End Sub

Private Sub Command14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Not MSComm1.PortOpen Then
        MSComm1.PortOpen = True
    End If
    KeyL = True
    KeyLeft
End Sub

Private Sub Command14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    KeyL = False
    MSComm1.PortOpen = False
End Sub

Private Sub Command15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Not MSComm1.PortOpen Then
        MSComm1.PortOpen = True
    End If
    KeyR = True
    KeyRight
End Sub

Private Sub Command15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    KeyR = False
    MSComm1.PortOpen = False
End Sub

Private Sub Command16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Not MSComm1.PortOpen Then
        MSComm1.PortOpen = True
    End If
    KeyD = True
    KeyDown
End Sub

Private Sub Command16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    KeyD = False
    MSComm1.PortOpen = False
End Sub

Private Sub MSComm1_OnComm()
On Error Resume Next
    'If MSComm1.CommEvent Then
    '    If (kxy = 0) Then
    '        Timer2.Enabled = True
    '    End If
    '    DataIn = MSComm1.Input
    '    Data = Asc(DataIn)
    '    kxy = kxy + 1
    '    DataPlot(kxy) = Data
    '    Data = (Data / 255) * 5
    '    DataInput(kxy) = Str(Data)
    '    Label7 = Format(Data, "00.00")
    '    LocationXY
    '    ProgressBar
    '    CCC = False
    '        If kxy >= xy Then
    '            MSComm1.Output = "e"
    '            Timer2.Enabled = False
    '            MSComm1.PortOpen = False
    '            Label5 = "00:00:00"
    '            ProgressBar1.Value = 100
    '            displayScanData.Hide
    '            Scan1.Visible = True
    '            Scan3.Visible = True
    '            Command3.Value = True
    '            Scan1.ZOrder
    '            SaveData
    '            ClearScan
    '        End If
    'End If
End Sub

Private Sub Option1_Click()
On Error Resume Next
    ComPortValue = 3
End Sub

Private Sub Option2_Click()
On Error Resume Next
    ComPortValue = 4
End Sub

Private Sub Option3_Click()
On Error Resume Next
    ComPortValue = 1
End Sub

Private Sub Option4_Click()
On Error Resume Next
    ComPortValue = 2
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If coppic = True Then
        If (X <= 9) And (X >= 0) Then
            x00 = X
        End If
        If (Y <= 9) And (Y >= 0) Then
            y00 = Y
        End If
    Label2(1) = "X-axis = " & Format(X, "0")
    Label2(0) = "Y-axis = " & Format(Y, "0")
        Cop
    End If
    Label2(1) = "X-axis = " & Format(X, "0")
    Label2(0) = "Y-axis = " & Format(Y, "0")
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    coppic = True
    If (X <= 9) And (X >= 0) Then
        x0 = X
    End If
    If (Y <= 9) And (Y >= 0) Then
        y0 = Y
    End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    coppic = False
    Cut_Area
    Command2.Enabled = True
End Sub

Sub Cop()
On Error Resume Next
   Do
    Picture2.Refresh
    Picture2.Line (x0, y0)-(x00, y00), vbRed, B
    DoEvents
   Loop While coppic
End Sub

Sub Cut_Area()
On Error Resume Next
    If x0 > x00 Then
        sx1 = x00
        sx2 = x0
    Else
        sx1 = x0
        sx2 = x00
    End If
    If y0 > y00 Then
        sy1 = y00
        sy2 = y0
    Else
        sy1 = y0
        sy2 = y00
    End If
    Picture2.AutoRedraw = True
    Picture2.Line (x0, y0)-(x00, y00), vbRed, B
    Picture2.AutoRedraw = False
    Combo1(3) = Format(sx1, "00.00")
    Combo1(4) = Format(sx2, "00.00")
    Combo1(5) = Format(sy1, "00.00")
    Combo1(6) = Format(sy2, "00.00")
End Sub

Sub DataToPort()
On Error Resume Next
        Dim xdatap As String
        MSComm1.Output = Chr$(&H1B)
        
        For i = 1 To 2 Step 1
            DataP = Val(Combo1(i)) * 100
             CutDataTo_DB9
        DoEvents
        Next i
        
        For i = 3 To 6 Step 1
            DataP = Val(Combo1(i))
            CutDataTo_DB9
        DoEvents
        Next i
End Sub

Sub CutDataTo_DB9()
 Dim cc As String
 
 On Error Resume Next
    cc = Format(DataP, "000")
    For i = 1 To 3 Step 1
     
     MSComm1.Output = Mid(cc, i, 1)
    Next
    
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If kxy <= 0 Then
        CountXY = CountXY + 1
        Scan3.Visible = True
        Text15(2) = "Y = " & Combo1(5) & " cm. And X = " & Combo1(3) & " cm."
        Scan3.ZOrder
            Select Case (CountXY Mod 2)
                Case 1
                    Shape1.BackColor = vbRed
                Case 0
                    Shape1.BackColor = vbBlue
            End Select
    Else
        Timer1.Enabled = False
        Scan3.Visible = False
        CountXY = 0
    End If
End Sub

Sub HHH1()
    On Error Resume Next
    Do
        MSComm1.Output = "h"
        DoEvents
    Loop While HHH
End Sub

Sub CCC1()
On Error Resume Next
    Do
        MSComm1.Output = "c"
        DoEvents
    Loop While CCC
End Sub

Sub KeyDown()
On Error Resume Next
    Do
    MSComm1.Output = "Z"
    DoEvents
    Loop While KeyD
End Sub

Sub KeyUp()
On Error Resume Next
    Do
    MSComm1.Output = "A"
    DoEvents
    Loop While KeyU
End Sub

Sub KeyRight()
On Error Resume Next
    Do
    MSComm1.Output = ">"
    DoEvents
    Loop While KeyR
End Sub

Sub KeyLeft()
On Error Resume Next
    Do
    MSComm1.Output = "<"
    DoEvents
    Loop While KeyL
End Sub

Sub E_Time1()
Dim tY As Currency
Dim tX As Currency
Dim tYstep As Currency
On Error Resume Next
       
    tY = (Val(Combo1(6)) * 10) - (Val(Combo1(5)) * 10)
    tX = (Val(Combo1(4)) * 10) - (Val(Combo1(3)) * 10)
    tYstep = Val(Combo1(2))
    tt2 = ((1 + (tY / tYstep)) * 0.5 * tX) + (tY * 1)
    
    hh = tt2 \ 3600
    hh1 = tt2 Mod 3600
    mm = hh1 \ 60
    ss = hh1 Mod 60
    ss = ss
    Label5 = Format(hh, "00") & ":" & Format(mm, "00") & ":" & Format(ss, "00")
End Sub

Sub E_Time2()
On Error Resume Next
    tt2 = tt2 - 1
    hh = tt2 \ 3600
    hh1 = tt2 Mod 3600
    mm = hh1 \ 60
    ss = hh1 Mod 60
    If (hh = 0) And (mm = 0) And (ss = 2) Then Timer2.Enabled = False
    Label5 = Format(hh, "00") & ":" & Format(mm, "00") & ":" & Format(ss, "00")
End Sub

Sub LocationXY()
On Error Resume Next
    lll = kxy
    If (lll Mod ax1) = 0 Then
        xxx1 = ax1
        yyy = (lll \ ax1)
    Else
        xxx1 = (lll Mod ax1)
        yyy = (lll \ ax1) + 1
    End If
    Select Case (yyy Mod 2)
        Case 1
            xxx = xxx1
        Case 0
            xxx = (ax1 - xxx1) + 1
    End Select
    xxx = Val(Combo1(3)) + ((xxx - 1) * ((Val(Combo1(1))) / 10))
    yyy = Val(Combo1(5)) + ((yyy - 1) * ((Val(Combo1(2))) / 10))
    Label3 = Format(yyy, "00.00") & " , " & Format(xxx, "00.00")
End Sub

Sub ProgressBar()
On Error Resume Next
        hun = (kxy / xy) * 100
        ProgressBar1.Value = hun
        Label9 = hun & "%"
End Sub

Sub SetComPort()
On Error Resume Next
    Select Case ComPortValue
        Case 1
            MSComm1.CommPort = 1
        Case 2
            MSComm1.CommPort = 2
        Case 3
            MSComm1.CommPort = 3
        Case 4
            MSComm1.CommPort = 4
    End Select
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.InputLen = 1
    MSComm1.RThreshold = 1
        If Not MSComm1.PortOpen Then
            MSComm1.PortOpen = True
        End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
    E_Time2
End Sub

Sub CheckInput()
On Error Resume Next
    If (Combo1(1) = "" Or Combo1(2) = "" Or Combo1(3) = "" Or Combo1(4) = "" Or Combo1(5) = "" Or Combo1(6) = "") Then
            Command2.Enabled = False
    Else
        CheckInput2
    End If
End Sub

Sub CheckInput2()
    If (CInt(Combo1(3)) > CInt(Combo1(4))) Then
        If MsgBox("Xmin > Xmax", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (CInt(Combo1(5)) > CInt(Combo1(6))) Then
        If MsgBox("Ymin > Ymax", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(1)) < 0.02) Then
        If MsgBox("Xscale < 0.02 mm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
 ElseIf (Val(Combo1(2)) < 0.02) Then
        If MsgBox("Yscale < 0.02 mm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(3)) < 0) Then
        If MsgBox("Xmin <  0 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(4)) < 0) Then
        If MsgBox("Xmax <  0 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(5)) < 0) Then
        If MsgBox("Ymin <  0 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
   ElseIf (Val(Combo1(6)) < 0) Then
        If MsgBox("Ymax <  0 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(3)) > 20) Then
        If MsgBox("Xmin >  20 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(4)) > 20) Then
        If MsgBox("Xmax >  20 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(5)) > 20) Then
        If MsgBox("Ymin >  20 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    ElseIf (Val(Combo1(6)) > 20) Then
        If MsgBox("Ymax >  20 cm.", vbOKOnly, "Data Input Error") = vbOK Then
            Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
    End If
End Sub

Sub ClearScan()
    Label3 = ""
    Label5 = ""
    Label7 = ""
    Label9 = ""
    ProgressBar1.Value = 0
    Timer1.Enabled = False
    Timer2.Enabled = False
    MSComm1.PortOpen = False
End Sub
