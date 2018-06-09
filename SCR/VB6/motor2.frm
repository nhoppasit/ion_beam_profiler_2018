VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   1320
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7245
   Begin VB.Frame Frame5 
      BackColor       =   &H80000000&
      Caption         =   "Set Positions of Motor"
      Height          =   1935
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   2895
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   615
         Index           =   2
         Left            =   960
         Picture         =   "motor2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   1335
         Index           =   0
         Left            =   240
         Picture         =   "motor2.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   1335
         Index           =   1
         Left            =   2040
         Picture         =   "motor2.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmd_mot_mov 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Height          =   615
         Index           =   3
         Left            =   960
         Picture         =   "motor2.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   600
      TabIndex        =   19
      Top             =   5400
      Width           =   1095
   End
   Begin VB.PictureBox pct_area 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   3840
      MousePointer    =   2  'Cross
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   18
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Area of  XY Axis"
      Height          =   1215
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   330
         Index           =   3
         ItemData        =   "motor2.frx":1108
         Left            =   960
         List            =   "motor2.frx":111B
         TabIndex        =   11
         Text            =   "5"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Index           =   4
         ItemData        =   "motor2.frx":1131
         Left            =   3000
         List            =   "motor2.frx":1144
         TabIndex        =   10
         Text            =   "20"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Index           =   5
         ItemData        =   "motor2.frx":115A
         Left            =   960
         List            =   "motor2.frx":116D
         TabIndex        =   9
         Text            =   "5"
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Index           =   6
         ItemData        =   "motor2.frx":1183
         Left            =   3000
         List            =   "motor2.frx":1196
         TabIndex        =   8
         Text            =   "20"
         Top             =   720
         Width           =   1455
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm."
         Height          =   210
         Index           =   9
         Left            =   4545
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Y axis"
         Height          =   210
         Index           =   10
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm."
         Height          =   210
         Index           =   11
         Left            =   4545
         TabIndex        =   14
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "to"
         Height          =   210
         Index           =   12
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "to"
         Height          =   210
         Index           =   13
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scale Scan of  XY Axis"
      Height          =   735
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   330
         Index           =   1
         ItemData        =   "motor2.frx":11AC
         Left            =   960
         List            =   "motor2.frx":11BF
         TabIndex        =   2
         Text            =   "0.10"
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Index           =   2
         ItemData        =   "motor2.frx":11E1
         Left            =   3360
         List            =   "motor2.frx":11F4
         TabIndex        =   1
         Text            =   "0.10"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm."
         Height          =   210
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "mm."
         Height          =   210
         Index           =   3
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "step x"
         Height          =   210
         Index           =   4
         Left            =   390
         TabIndex        =   4
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "step y"
         Height          =   210
         Index           =   5
         Left            =   2790
         TabIndex        =   3
         Top             =   360
         Width           =   435
      End
   End
   Begin MSCommLib.MSComm Serial 
      Left            =   6120
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const data_port = &H378
Private Const control_port = &H37A
Private Const x_axis = &H1
Private Const y_axis = &H2
Private Const z_axis = &H4
'unsigned int mot_step, mot_speed;
'unsigned int x_step,y_step;
'unsigned int x_min,x_max;
'unsigned int y_min,y_max;

Private Const x_left = 1

Private x_cord1 As Integer
Private y_cord1 As Integer
Private x_cord2 As Integer
Private y_cord2 As Integer
Private DataP As Integer
Private mot_flag As Boolean
'Dim a As Integer


'Private btn_flag As Boolean

Private Sub motor_move(axis As Integer)
    Dim direction As String * 1
    If Not Serial.PortOpen Then
        Serial.PortOpen = True
    End If
    mot_flag = True
    
    Select Case axis
        Case 0: direction = "<"
        Case 1: direction = ">"
        Case 2: direction = "A"
        Case 3: direction = "Z"
    End Select
    Do
        Serial.Output = direction
        DoEvents
    Loop While mot_flag
End Sub

Private Sub Send_Data(sData As String)
    'DataP = Serial.Input
    Serial.Output = sData
    Do
     DoEvents
    Loop Until (Serial.Input = "s")
    
End Sub
Private Sub SetComPort()
  'On Error Resume Next
    Serial.CommPort = 1
    Serial.Settings = "9600,n,8,1"
    Serial.InputLen = 1
    Serial.RThreshold = 1
    If Not Serial.PortOpen Then
       Serial.PortOpen = True
    End If
End Sub

Private Sub DataToPort()
  On Error Resume Next
     Dim xdatap As String
     Dim i As Integer
     'Serial.Output = Chr$(&H1B)
     Send_Data (Chr$(&H1B))
        For i = 1 To 2 Step 1
            DataP = Val(Combo1(i)) * 100
             CutDataTo_DB9
        'DoEvents
        Next i
        
        For i = 3 To 6 Step 1
            DataP = Val(Combo1(i))
            CutDataTo_DB9
        'DoEvents
        Next i
End Sub

Sub CutDataTo_DB9()
 Dim cc As String
 Dim i As Integer
 
 On Error Resume Next
    cc = Format(DataP, "000")
    For i = 1 To 3 Step 1
     
     'Serial.Output = Mid(cc, i, 1)
     Send_Data (Mid(cc, i, 1))
    Next
    
End Sub

Private Sub cmd_mot_mov_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   motor_move (Index)
End Sub

Private Sub cmd_mot_mov_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mot_flag = False
End Sub

Private Sub Command1_Click()
    SetComPort
    DataToPort
End Sub

Private Sub Command14_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call siclcleanup    ' Tell SICL to clean up for this task
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
