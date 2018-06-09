VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_main 
   Caption         =   "  Ion Beam Profiler v 1.0J"
   ClientHeight    =   6390
   ClientLeft      =   810
   ClientTop       =   2565
   ClientWidth     =   12360
   Icon            =   "x-y motor3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   12360
   Begin VB.Frame Frame7 
      Caption         =   "y - axis"
      Height          =   855
      Left            =   5760
      TabIndex        =   67
      Top             =   4320
      Width           =   2415
      Begin VB.CommandButton cmd_perset_y_axis 
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
         Left            =   1440
         TabIndex        =   69
         ToolTipText     =   "Preset to a position"
         Top             =   380
         Width           =   735
      End
      Begin VB.ComboBox cbo_y_post 
         Height          =   330
         ItemData        =   "x-y motor3.frx":0ABA
         Left            =   120
         List            =   "x-y motor3.frx":0ADC
         TabIndex        =   68
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Index           =   7
         Left            =   1080
         TabIndex        =   70
         Top             =   375
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "x - axis"
      Height          =   855
      Left            =   5760
      TabIndex        =   63
      Top             =   3360
      Width           =   2415
      Begin VB.ComboBox cbo_x_post 
         Height          =   330
         ItemData        =   "x-y motor3.frx":0B07
         Left            =   120
         List            =   "x-y motor3.frx":0B29
         TabIndex        =   65
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmd_perset_x_axis 
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
         Left            =   1440
         TabIndex        =   64
         ToolTipText     =   "Preset to a position"
         Top             =   380
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Index           =   6
         Left            =   1080
         TabIndex        =   66
         Top             =   375
         Width           =   255
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "SAMPLE RATE"
      Height          =   1215
      Left            =   4080
      TabIndex        =   53
      Top             =   1080
      Width           =   1335
      Begin VB.ComboBox cbo_sample_data_rate 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         ItemData        =   "x-y motor3.frx":0B54
         Left            =   240
         List            =   "x-y motor3.frx":0B76
         TabIndex        =   54
         Text            =   "2"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "sec/point"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   55
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame10 
      Height          =   855
      Left            =   120
      TabIndex        =   48
      Top             =   5400
      Width           =   5295
      Begin VB.CommandButton cmd_save 
         Caption         =   "&SAVE"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1080
         TabIndex        =   57
         ToolTipText     =   "Save Data"
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton cmd_run 
         Caption         =   "&RUN"
         Height          =   490
         Left            =   120
         TabIndex        =   52
         ToolTipText     =   "Start Running Scan"
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "S&TOP"
         Enabled         =   0   'False
         Height          =   490
         Left            =   2040
         MousePointer    =   1  'Arrow
         TabIndex        =   51
         ToolTipText     =   "Stop Running Scan"
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton cmd_setup 
         Caption         =   "SETUP >>"
         Height          =   490
         Left            =   4080
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Setup Parameters"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "E&XIT"
         Height          =   490
         Left            =   3120
         TabIndex        =   49
         ToolTipText     =   "Quit Application"
         Top             =   240
         Width           =   800
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2775
      Left            =   2520
      TabIndex        =   32
      Top             =   2520
      Width           =   2895
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
         Index           =   5
         Left            =   2000
         TabIndex        =   62
         ToolTipText     =   "press SHIFT button & Cluick mouse"
         Top             =   1440
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
         Index           =   4
         Left            =   1080
         TabIndex        =   61
         ToolTipText     =   "press SHIFT button & Cluick mouse"
         Top             =   1440
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
         Index           =   3
         Left            =   120
         TabIndex        =   60
         ToolTipText     =   "press SHIFT button & Click mouse"
         Top             =   1440
         Width           =   550
      End
      Begin VB.Frame Frame9 
         Caption         =   "z - axis"
         Height          =   855
         Left            =   0
         TabIndex        =   45
         Top             =   1920
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
            TabIndex        =   47
            ToolTipText     =   "Preset to a position"
            Top             =   380
            Width           =   735
         End
         Begin VB.ComboBox cbo_z_post 
            Height          =   330
            ItemData        =   "x-y motor3.frx":0B99
            Left            =   240
            List            =   "x-y motor3.frx":0BBB
            TabIndex        =   46
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
            TabIndex        =   56
            Top             =   380
            Width           =   255
         End
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "ORG"
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
         TabIndex        =   35
         ToolTipText     =   "Goto Origin X position"
         Top             =   1080
         Width           =   550
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "ORG"
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
         Left            =   1080
         TabIndex        =   34
         ToolTipText     =   "Goto Origin Y position"
         Top             =   1080
         Width           =   550
      End
      Begin VB.CommandButton cmd_reset 
         Caption         =   "ORG"
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
         TabIndex        =   33
         ToolTipText     =   "Goto Origin Z position"
         Top             =   1080
         Width           =   550
      End
      Begin VB.Label lbl_x_post 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   100
         TabIndex        =   44
         Top             =   600
         Width           =   550
      End
      Begin VB.Label lbl_y_post 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1050
         TabIndex        =   43
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lbl_z_post 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2000
         TabIndex        =   42
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "x-position"
         Height          =   300
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "y-position"
         Height          =   300
         Left            =   1080
         TabIndex        =   40
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "z-position"
         Height          =   300
         Left            =   2040
         TabIndex        =   39
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Left            =   705
         TabIndex        =   38
         Top             =   620
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Left            =   2600
         TabIndex        =   37
         Top             =   620
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         Height          =   210
         Left            =   1650
         TabIndex        =   36
         Top             =   620
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cmdlog1 
      Left            =   0
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "SETUP"
      Height          =   3135
      Left            =   5760
      TabIndex        =   25
      Top             =   120
      Width           =   2415
      Begin VB.Frame Frame6 
         Caption         =   "INSTUMENT ADDR"
         Height          =   1815
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1935
         Begin VB.ComboBox cbo_dmm_addr 
            Height          =   330
            Index           =   1
            ItemData        =   "x-y motor3.frx":0BE6
            Left            =   885
            List            =   "x-y motor3.frx":0C47
            TabIndex        =   58
            Text            =   "24"
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cbo_dmm_addr 
            Height          =   330
            Index           =   0
            ItemData        =   "x-y motor3.frx":0CBD
            Left            =   840
            List            =   "x-y motor3.frx":0D1E
            TabIndex        =   30
            Text            =   "24"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "DMM2"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   59
            Top             =   1125
            Width           =   810
         End
         Begin VB.Label Label10 
            Caption         =   "DMM1"
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   31
            Top             =   525
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "HPIB ADDR"
         Height          =   855
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1935
         Begin VB.ComboBox cbo_hpib_addr 
            Height          =   330
            ItemData        =   "x-y motor3.frx":0D94
            Left            =   840
            List            =   "x-y motor3.frx":0DC5
            TabIndex        =   27
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "HPIB"
            Height          =   255
            Left            =   195
            TabIndex        =   28
            Top             =   400
            Width           =   810
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000000&
      Caption         =   "Set Positions of Motor"
      Height          =   2775
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   5295
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   495
         Index           =   4
         Left            =   150
         Picture         =   "x-y motor3.frx":0DFB
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "move forward"
         Top             =   2160
         Width           =   900
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   495
         Index           =   5
         Left            =   1400
         Picture         =   "x-y motor3.frx":123D
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "move backward"
         Top             =   2160
         Width           =   900
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   630
         Index           =   2
         Left            =   840
         Picture         =   "x-y motor3.frx":167F
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "move up"
         Top             =   360
         Width           =   750
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   1335
         Index           =   0
         Left            =   120
         Picture         =   "x-y motor3.frx":1AC1
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "x-y motor3.frx":1F03
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "x-y motor3.frx":2345
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "move down"
         Top             =   1080
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCAN AERA"
      Height          =   1215
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3855
      Begin VB.ComboBox cbo_x_min 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor3.frx":2787
         Left            =   960
         List            =   "x-y motor3.frx":279A
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbo_x_max 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor3.frx":27B0
         Left            =   2280
         List            =   "x-y motor3.frx":27C3
         TabIndex        =   10
         Text            =   "20"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbo_y_min 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor3.frx":27D9
         Left            =   960
         List            =   "x-y motor3.frx":27EC
         TabIndex        =   9
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cbo_y_max 
         BackColor       =   &H00FFFFC0&
         Height          =   330
         ItemData        =   "x-y motor3.frx":2802
         Left            =   2280
         List            =   "x-y motor3.frx":2815
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
      Caption         =   "SCAN RESOLUTION"
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox cbo_x_step 
         BackColor       =   &H00C0C0FF&
         Height          =   330
         ItemData        =   "x-y motor3.frx":282B
         Left            =   960
         List            =   "x-y motor3.frx":284A
         TabIndex        =   2
         Text            =   "0.02"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cbo_y_step 
         BackColor       =   &H00C0C0FF&
         Height          =   330
         ItemData        =   "x-y motor3.frx":2884
         Left            =   3600
         List            =   "x-y motor3.frx":28A3
         TabIndex        =   1
         Text            =   "0.02"
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
   Begin VB.Frame fmAllLimit 
      Caption         =   "Limit"
      Height          =   2775
      Left            =   8400
      TabIndex        =   71
      Top             =   120
      Width           =   3375
      Begin VB.Frame fmZLimit 
         Caption         =   "Z Limit"
         Height          =   735
         Left            =   240
         TabIndex        =   82
         Top             =   1920
         Width           =   2895
         Begin VB.TextBox txtZMinLimit 
            Height          =   315
            Left            =   480
            TabIndex        =   84
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtZMaxLimit 
            Height          =   315
            Left            =   1920
            TabIndex        =   83
            Text            =   "90"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Min"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "Max"
            Height          =   255
            Left            =   1560
            TabIndex        =   85
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "X Limit"
         Height          =   735
         Left            =   240
         TabIndex        =   77
         Top             =   240
         Width           =   2895
         Begin VB.TextBox txtXMinLimit 
            Height          =   315
            Left            =   480
            TabIndex        =   79
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtXMaxLimit 
            Height          =   315
            Left            =   1920
            TabIndex        =   78
            Text            =   "23"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Min"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label14 
            Caption         =   "Max"
            Height          =   255
            Left            =   1560
            TabIndex        =   80
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Y Limit"
         Height          =   735
         Left            =   240
         TabIndex        =   72
         Top             =   1080
         Width           =   2895
         Begin VB.TextBox txtYMinLimit 
            Height          =   315
            Left            =   480
            TabIndex        =   74
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtYMaxLimit 
            Height          =   315
            Left            =   1920
            TabIndex        =   73
            Text            =   "23"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Min"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "Max"
            Height          =   255
            Left            =   1560
            TabIndex        =   75
            Top             =   240
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private xlApp As Object     'excel.Application
Private xlBook As Object     'excel.workbooks
Private xlSheet As Object   'excel.WorkSheet

Private flag_plot As Boolean
Private motor_flag As Boolean


Private hpib_dmm As String
Private data_sample_count As String
Private data_sample_rate As String

Private data_port As Integer
Private control_port As Integer

Private Const x_axis = 4 '&HCF          'for ETT-82545 card
Private Const y_axis = 5 '&HCE
Private Const z_axis = 6 '&HCD
Private Const control8255 = 7 '&HC4

Private Const wr = 2   'Pin 16

Private x_post As Single
Private y_post As Single
Private z_post As Single

Private step_fwd(0 To 9) As Byte
Private step_rwd(0 To 9) As Byte
Private m_speed As Integer
Private m_step As Integer

Private id As Integer     ' device session id

Private Const x_left = 1

Private x_cord1 As Integer
Private y_cord1 As Integer
Private x_cord2 As Integer
Private y_cord2 As Integer
'Private DataP As Integer


Dim x_max As Double
Dim y_max As Double
Dim data_read_xy() As Single





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
        .Filter = "Text Files(*.txt)|*.txt|Ion Source Files(*.ibp)|*.ibp"
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
  
  'hpib_dmm = "hpib7,7" '"gpib" & cbo_hpib_addr.Text & "," & cbo_dmm_addr(0).Text
  
  data_sample_count = "SAMP:COUN 1" & vbLf '& cbo_sample_data_point.Text + Chr$(10)
  data_sample_rate = "TRIG:DELAY " & cbo_sample_data_rate.Text & vbLf
End Sub



Private Function read_dmm() As String
                  
   Dim readbuf As String * 200      ' buffer used for iread
   Dim nargs As Integer             ' # args converted by format string

'  Set up an error handler within this subroutine that will get
'  called if a SICL error occurs.
   On Error GoTo ErrorHandler

   'nargs = ivprintf(id, "READ?" + Chr$(10))
   nargs = ivprintf(DMM_ID, "INIT" + Chr$(10))
   nargs = ivprintf(DMM_ID, "*TRG" + Chr$(10))
   nargs = ivprintf(DMM_ID, "FETC?" + Chr$(10))
   
   nargs = ivscanf(DMM_ID, "%200t", readbuf)

'  Close the device session.
 '  Call iclose(id)
   read_dmm = readbuf
   Exit Function

ErrorHandler:
'  Close the device session if iopen was successful.
   If id <> 0 Then
      iclose (DMM_ID)
   End If
   MsgBox Error$, vbOKOnly, "ERROR"
   read_dmm = "0"
   Exit Function
End Function

Private Function initialize_dmm() As Boolean
                  
   'Dim readbuf As String * 200      ' buffer used for iread
   'Dim commandstr As String * 128   ' command passed to instrument
   'Dim index As Integer             ' used to parse SCPI error message
   Dim nargs As Integer             ' # args converted by format string

'  Set up an error handler within this subroutine that will get
'  called if a SICL error occurs.
   Call get_setup_parameter
   
   Initialize_Instrument DMM_ID, DMM_AD

'  Open a device session using the device address contained in
   'If id = 0 Then
    'id = iopen(hpib_dmm)
   'End If
   'Me.Caption = id
'  Set the I/O timeout value for this session to 1 second.
   Call itimeout(DMM_ID, 10000)
    
'  Clear the error/event queue for the instrument.  This allows
'  us to query the instrument after sending a command to see if
'  the command was accepted.
   nargs = ivprintf(DMM_ID, "*CLS" & vbLf)
   
   nargs = ivprintf(DMM_ID, "CONF:volt:DC 2,0.0001" + Chr$(10)) ' uncomment this line for voltage
   'nargs = ivprintf(DMM_ID, "volt:DC:nplc 10" + Chr$(10))
   
   'nargs = ivprintf(DMM_ID, "CONF:CURR:DC 0.00001,0.00001" & vbLf)    ' uncomment this line for current
   'nargs = ivprintf(id, "CURR:DC:nplc 10" & vbLf)                  ' uncomment this line for current
   
   nargs = ivprintf(DMM_ID, "TRIG:SOUR BUS" & vbLf)
   nargs = ivprintf(DMM_ID, data_sample_count)
      
   nargs = ivprintf(DMM_ID, data_sample_rate)               'delay 2 sec for volt   ---- 0 for current
      
   'nargs = ivprintf(id, "TRIG:COUN 1" + Chr$(10))
   'nargs = ivprintf(id, "*CLS" + Chr$(10))
   

'  Close the device session.
 '  Call iclose(id)
   initialize_dmm = True
   
End Function

Private Sub scan(ByRef data_read_xy() As Single, x_step As Integer, x_max As Integer, _
                 y_step As Integer, y_max As Integer)
  
  
  Static Sheet
  
  Dim i As Integer
  Dim j As Integer
  Dim data_read_xy_s() As String * 200

  Dim data_x As String
  
  If x_max > 256 Then
    MsgBox "number of row > 256"
    Exit Sub
  End If
  
  ReDim data_read_xy_s(0 To y_max, 0 To x_max)

  
  On Error GoTo errtrap
 
  If Sheet = 0 Then
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.workbooks.Add
  End If
    'Make Excel visible through the Application object only once
  xlApp.Visible = True

  Sheet = Sheet + 1

  If Sheet < 4 Then
    Set xlSheet = xlBook.worksheets(Sheet)
  Else
    Set xlSheet = xlBook.worksheets.Add
  End If
  
  xlSheet.Select
  ' Set Name of Sheet
  xlSheet.name = "Focus Ion Beam" & Str$(Sheet)
  
  flag_plot = True
  
  Me.MousePointer = 11
  
  For j = 0 To y_max
   
   For i = 0 To x_max - 1
      If (j Mod 2) = 0 Then
        
        data_read_xy_s(j, i) = read_dmm()
        xlSheet.Cells(j + 1, i + 1).value = str2sng(data_read_xy_s(j, i))
        
        Write_Instrument MMC_2XY_ID, "M:XP" & x_step, vbCrLf
        Write_Instrument MMC_2XY_ID, "G:", vbCrLf
        
        update_xy_pst
      Else
        data_read_xy_s(j, x_max - i) = read_dmm()
        xlSheet.Cells(j + 1, x_max - i + 1).value = str2sng(data_read_xy_s(j, x_max - i))
        
        Write_Instrument MMC_2XY_ID, "M:XP-" & x_step, vbCrLf
        Write_Instrument MMC_2XY_ID, "G:", vbCrLf
        
        update_xy_pst
      End If
      
      If Not flag_plot Then GoTo abort_scan
     
      DoEvents
   Next i
   
   If (j Mod 2) = 0 Then
     data_read_xy_s(j, x_max) = read_dmm()
     xlSheet.Cells(j + 1, x_max + 1).value = str2sng(data_read_xy_s(j, x_max))
   Else
     data_read_xy_s(j, 0) = read_dmm()
     xlSheet.Cells(j + 1, 1).value = str2sng(data_read_xy_s(j, 0))
   End If
     
   If j <> y_max Then
     Write_Instrument MMC_2XY_ID, "M:YP" & y_step, vbCrLf
     Write_Instrument MMC_2XY_ID, "G:", vbCrLf
     update_xy_pst
   End If
  Next j
  
abort_scan: 'CALL WRITE FILE
  
  For j = 0 To y_max
    For i = 0 To x_max
      data_read_xy(j, i) = str2sng(data_read_xy_s(j, i))
    Next i
  Next j
  
  Me.MousePointer = 0
  
  Exit Sub
  
errtrap:
  If Err.Number = -2147417848 Then
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.workbooks.Add
    xlApp.Visible = True
    Sheet = 1
    'If Sheet < 4 Then
    Set xlSheet = xlBook.worksheets(Sheet)
   
   'Else
    Resume Next
  ElseIf Err.Number <> 0 Then
    MsgBox Err.Number & "  " & Err.Description & "  " & Err.Source
  End If
  
End Sub

Private Sub preset_scan(x_mn As Double, y_mn As Double)
  Dim x_step As String
  Dim y_step As String
  
  x_step = Str$(CInt(x_mn * 500))
  y_step = Str$(CInt(y_mn * 500))
  
  Write_Instrument MMC_2XY_ID, "A:WP" & x_step & "P" & y_step, vbCrLf
  update_xy_pst
End Sub

Private Sub scan_data()
 Dim x_step As Double
 Dim y_step As Double
 Dim x_min As Double
 Dim y_min As Double
 Dim x_taltal_step As Integer
 Dim y_taltal_step As Integer
 
 x_step = Val(cbo_x_step.Text) * 500        'convert distance to step(1 mm = 500 step)
 y_step = Val(cbo_y_step.Text) * 500
 x_min = Val(cbo_x_min.Text)
 x_max = Val(cbo_x_max.Text)
 y_min = Val(cbo_y_min.Text)
 y_max = Val(cbo_y_max.Text)
 
 x_taltal_step = CInt((x_max - x_min) * 500 / x_step)
 y_taltal_step = CInt((y_max - y_min) * 500 / y_step)
 
 'x_taltal_step = CInt(x_max)
 'y_taltal_step = CInt(y_max)
 
 Call preset_scan(x_min, y_min)
 ReDim data_read_xy(0 To y_taltal_step, 0 To x_taltal_step)
 
 Call scan(data_read_xy(), CInt(x_step), x_taltal_step, CInt(y_step), y_taltal_step)
 Call save_data_file(data_read_xy(), x_taltal_step, y_taltal_step)

End Sub
Private Sub xyz_move(motor_phase() As Byte, motor_axis As Byte, _
            motor_Step As Integer, motor_speed)
  
  Dim i As Integer
  
  For i = 0 To motor_Step - 1
      'Out data_port, motor_phase(i Mod 10)
      'Out control_port, motor_axis
      
      'ClrPortBit control_port, wr
      'SetPortBit control_port, wr
      
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
 'Call motor_off
 
 x_post = 0
 y_post = 0
 z_post = 0
  
 'Call update_xyz_post(x_post, 0, lbl_x_post)
 'Call update_xyz_post(y_post, 0, lbl_y_post)
 'Call update_xyz_post(z_post, 0, lbl_z_post)
End Sub
Private Sub motor_move(axis As Integer)
    
    Dim motor_axis As String * 1
    Dim motor_Step As String
    Dim mm_of_50P As Double
    Dim dXPos As Double
    Dim dYPos As Double
    Dim dZPos As Double
    Dim end_post As String * 1
    Dim post As String
    
    motor_flag = True

    
    Select Case axis
        Case 0:  motor_axis = "X"
                 motor_Step = "P-50"              'move left
        Case 1:  motor_axis = "X"
                 motor_Step = "P50"               'move right"
        Case 2:  motor_axis = "Y"
                 motor_Step = "P-50"              'move up
        Case 3:  motor_axis = "Y"
                 motor_Step = "P50"               'move down
        Case 4:  motor_axis = "X"
                 motor_Step = "P-50"               'move +z
        Case 5:  motor_axis = "X"
                 motor_Step = "P50"               'move +z
                 
    End Select
    
    
    Do While motor_flag
        
        If axis < 4 Then
               
             Write_Instrument MMC_2XY_ID, "Q:", vbCrLf
             post = Read_Instrument(MMC_2XY_ID)
             end_post = Right$(post, 1)
             
             dXPos = Val(Left$(post, 9)) / 500
             lbl_x_post.Caption = Format(dXPos, "#0.000")
             
             dYPos = Val(Mid$(post, 11, 9)) / 500
             lbl_y_post.Caption = Format(dYPos, "#0.000")
  
             mm_of_50P = 50 / 500
             If axis = 1 Or axis = 3 Then
                dXPos = dXPos + mm_of_50P
                dYPos = dYPos + mm_of_50P
             Else
                dXPos = dXPos - mm_of_50P
                dYPos = dYPos - mm_of_50P
             End If
             
             If motor_axis = "X" Then
                
                If (dXPos <= Val(txtXMaxLimit.Text)) And _
                   (dXPos >= Val(txtXMinLimit.Text)) _
                Then
                   
                   Write_Instrument MMC_2XY_ID, "M:" & motor_axis & motor_Step, vbCrLf
                   Write_Instrument MMC_2XY_ID, "G:", vbCrLf
                  
                End If

             Else
             
                If (dYPos <= Val(txtYMaxLimit.Text)) And _
                   (dYPos >= Val(txtYMinLimit.Text)) _
                Then
                   
                   Write_Instrument MMC_2XY_ID, "M:" & motor_axis & motor_Step, vbCrLf
                   Write_Instrument MMC_2XY_ID, "G:", vbCrLf
                  
                End If
             
             End If
            
        Else
             '============================
             Write_Instrument MMC_2Z_ID, "Q:", vbCrLf
             post = Read_Instrument(MMC_2Z_ID)
             end_post = Right$(post, 1)
             
             dZPos = Val(Left$(post, 9)) / 500
             lbl_z_post.Caption = Format(dZPos, "#0.000")

             mm_of_50P = 50 / 500
             If axis = 5 Then
                dZPos = dZPos + mm_of_50P
             Else
                dZPos = dZPos - mm_of_50P
             End If
             If (dZPos <= Val(txtZMaxLimit.Text)) And _
                (dZPos >= Val(txtZMinLimit.Text)) _
             Then
                
                Write_Instrument MMC_2Z_ID, "M:" & motor_axis & motor_Step, vbCrLf
                Write_Instrument MMC_2Z_ID, "G:", vbCrLf
               
             End If
             '============================
           
        End If
                
        DoEvents
        
    Loop
    
End Sub

Private Sub cmd_exit_Click()
  Unload Me
End Sub

Private Sub cmd_mot_mov_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    motor_move (Index)
End Sub

Private Sub cmd_mot_mov_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    motor_flag = False
    update_xy_pst
    update_z_pst
End Sub

Private Sub cmd_perset_x_axis_Click()
    Dim pst As String
    
    '============================
    
    If (Val(cbo_x_post.Text) <= Val(txtXMaxLimit.Text)) And _
       (Val(cbo_x_post.Text) >= Val(txtXMinLimit.Text)) _
    Then
       
        pst = Str$(cbo_x_post.Text * 500)
        Write_Instrument MMC_2XY_ID, "A:XP" & pst, vbCrLf
      
    Else
        
        If Val(cbo_x_post.Text) > Val(lbl_x_post.Caption) Then
           cbo_x_post.Text = txtXMaxLimit.Text
        Else
           cbo_x_post.Text = txtXMinLimit.Text
        End If
       
        pst = Str$(cbo_x_post.Text * 500)
        Write_Instrument MMC_2XY_ID, "A:XP" & pst, vbCrLf
      
    End If
    
    update_xy_pst
End Sub

Private Sub cmd_perset_y_axis_Click()
    Dim pst As String
    
    '============================
    
    If (Val(cbo_y_post.Text) <= Val(txtYMaxLimit.Text)) And _
       (Val(cbo_y_post.Text) >= Val(txtYMinLimit.Text)) _
    Then
       
        pst = Str$(cbo_y_post.Text * 500)
        Write_Instrument MMC_2XY_ID, "A:YP" & pst, vbCrLf
      
    Else
        
        If Val(cbo_y_post.Text) > Val(lbl_y_post.Caption) Then
           cbo_y_post.Text = txtYMaxLimit.Text
        Else
           cbo_y_post.Text = txtYMinLimit.Text
        End If
       
        pst = Str$(cbo_y_post.Text * 500)
        Write_Instrument MMC_2XY_ID, "A:YP" & pst, vbCrLf
      
    End If
    
    '============================
    
    update_xy_pst
End Sub

Private Sub cmd_perset_z_axis_Click()
    Dim pst As String
    
    '============================
    
    If (Val(cbo_z_post.Text) <= Val(txtZMaxLimit.Text)) And _
       (Val(cbo_z_post.Text) >= Val(txtZMinLimit.Text)) _
    Then
       
        pst = Str$(cbo_z_post.Text * 500)
        Write_Instrument MMC_2Z_ID, "A:XP" & pst, vbCrLf
      
    Else
        
        If Val(cbo_z_post.Text) > Val(lbl_z_post.Caption) Then
           cbo_z_post.Text = txtZMaxLimit.Text
        Else
           cbo_z_post.Text = txtZMinLimit.Text
        End If
       
        pst = Str$(cbo_z_post.Text * 500)
        Write_Instrument MMC_2Z_ID, "A:XP" & pst, vbCrLf
      
    End If
    
    update_z_pst
End Sub



Private Sub cmd_reset_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
  Dim idx As Integer
    Dim pst As String
    
  idx = Index
    
  If Shift = 1 Then
    If MsgBox("Reset?", vbQuestion + vbOKCancel) <> vbOK Then Exit Sub
  
    Select Case idx
        Case 3: Write_Instrument MMC_2XY_ID, "P:14P0", vbCrLf
        Case 4: Write_Instrument MMC_2XY_ID, "P:15P0", vbCrLf
        Case 5: Write_Instrument MMC_2Z_ID, "P:14P0", vbCrLf
    End Select
    MsgBox "Reset OK"
    
  Else
   
    Select Case idx
        Case 0: Write_Instrument MMC_2XY_ID, "A:XP0", vbCrLf
        Case 1: Write_Instrument MMC_2XY_ID, "A:YP0", vbCrLf
        Case 2:
              '============================
              
              If (0 >= Val(txtZMinLimit.Text)) Then
                 
                  Write_Instrument MMC_2Z_ID, "A:XP0", vbCrLf
                
              Else
                  lbl_z_post.Caption = txtZMinLimit.Text
                  pst = Str$(txtZMinLimit.Text * 500)
                 
                  Write_Instrument MMC_2Z_ID, "A:XP" & pst, vbCrLf
                
              End If
            '=================================
       
            'Write_Instrument MMC_2Z_ID, "A:XP0", vbCrLf
   
    End Select
    
  End If
    
  update_xy_pst
  update_z_pst
  
    
End Sub

Private Sub cmd_run_Click()
  If Not initialize_dmm Then
   'MsgBox "NOT OK"
   Exit Sub
  End If
  'MsgBox "Scan Now"
  cmd_run.Enabled = False
  cmd_save.Enabled = False
  cmd_exit.Enabled = False
  cmd_cancel.Enabled = True
  Call scan_data
  cmd_run.Enabled = True
  cmd_save.Enabled = True
  cmd_exit.Enabled = True
  cmd_cancel.Enabled = False
  'Call motor_off
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

Private Sub cmd_save_Click()
    Call save_data_file(data_read_xy(), CInt(x_max), CInt(y_max))
End Sub

Private Sub cmd_setup_Click()
  If cmd_setup.Caption = "SETUP >>" Then
    cmd_setup.Caption = "SETUP <<"
    Me.width = 12030
  Else
    cmd_setup.Caption = "SETUP >>"
    Me.width = 5670
  End If
End Sub







Private Sub Form_DblClick()
    MsgBox "Ion Beam Profiler V 0.2a  (09-09-2003)" & vbCrLf _
         & "  by  Laser & Surface Physics Lab.   " & vbCrLf _
         & " Applied Physics Department, KMIT'L   ", vbInformation, "About Software"
End Sub

Private Sub Form_Load()
  If Not Initialize_Instrument(MMC_2XY_ID, MMC_2_XY_AD) Then Exit Sub
  If Not Initialize_Instrument(MMC_2Z_ID, MMC_2_Z_AD) Then Exit Sub
  update_xy_pst
  update_z_pst
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not cmd_run.Enabled Then
    MsgBox "Stop running first!", vbInformation
    Cancel = 1
  ElseIf MsgBox("Exit Program ?", vbOKCancel) = vbCancel Then
    Cancel = 1
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbNormal Then
   Me.Move (Screen.width - Me.width) / 2, (Screen.Height - Me.Height) / 2
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Finalized_Instrument MMC_2XY_ID
   Finalized_Instrument MMC_2Z_ID
   Call siclcleanup    ' Tell SICL to clean up for this task
   End
End Sub
Private Sub update_xy_pst()
  Dim end_post As String * 1
  Dim post As String
  Do
    Write_Instrument MMC_2XY_ID, "Q:", vbCrLf
    post = Read_Instrument(MMC_2XY_ID)
    end_post = Right$(post, 1)
    DoEvents
  Loop Until end_post = "R"
  'lbl_x_p.Caption = Left$(post, 9)
  lbl_x_post.Caption = Format(Val(Left$(post, 9)) / 500, "#0.000")
   
  'lbl_y_p.Caption = Mid$(post, 11, 9)
  lbl_y_post.Caption = Format(Val(Mid$(post, 11, 9)) / 500, "#0.000")
  
End Sub
Private Sub update_z_pst()
  Dim end_post As String * 1
  Dim post As String
  Do
    Write_Instrument MMC_2Z_ID, "Q:", vbCrLf
    post = Read_Instrument(MMC_2Z_ID)
    end_post = Right$(post, 1)
    DoEvents
  Loop Until end_post = "R"
  'lbl_x_p.Caption = Left$(post, 9)
  lbl_z_post.Caption = Format(Val(Left$(post, 9)) / 500, "#0.000")
   
 
End Sub

Private Sub txtXMaxLimit_LostFocus()
    If Val(txtXMinLimit.Text) >= Val(txtXMaxLimit.Text) Then
        txtXMinLimit.Text = "0"
        txtXMaxLimit.Text = "23"
    End If
End Sub

Private Sub txtXMinLimit_LostFocus()
    If Val(txtXMinLimit.Text) >= Val(txtXMaxLimit.Text) Then
        txtXMinLimit.Text = "0"
        txtXMaxLimit.Text = "23"
    End If
End Sub

Private Sub txtYMaxLimit_LostFocus()
    If Val(txtYMinLimit.Text) >= Val(txtYMaxLimit.Text) Then
        txtYMinLimit.Text = "0"
        txtYMaxLimit.Text = "23"
    End If
End Sub

Private Sub txtYMinLimit_LostFocus()
    If Val(txtYMinLimit.Text) >= Val(txtYMaxLimit.Text) Then
        txtYMinLimit.Text = "0"
        txtYMaxLimit.Text = "23"
    End If
End Sub

Private Sub txtZMaxLimit_LostFocus()
    If Val(txtZMinLimit.Text) >= Val(txtZMaxLimit.Text) Then
        txtZMinLimit.Text = "0"
        txtZMaxLimit.Text = "90"
    End If
End Sub

Private Sub txtZMinLimit_LostFocus()
    If Val(txtZMinLimit.Text) >= Val(txtZMaxLimit.Text) Then
        txtZMinLimit.Text = "0"
        txtZMaxLimit.Text = "90"
    End If
End Sub
