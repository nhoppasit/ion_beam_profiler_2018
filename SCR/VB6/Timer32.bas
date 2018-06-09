Attribute VB_Name = "bas_VBRC_Timer32"
Option Explicit
'
' Project Name  : -
' Library Name  : Benchmark Timer functions for VB/Win 5.0, 6.0
' VB/Win Module : TIMER32.BAS
' DLL Filename  : TIMER32.DLL
' Copyright     : (c)1996-1999 by Sutthisak Phongthanapanich
'

'
' TIMER32.DLL constants declaration.
'
' Copyright string character case.
Public Const BTR_COPYRIGHT_DEFAULT = 0
Public Const BTR_COPYRIGHT_CAPITAL = 1
Public Const BTR_COPYRIGHT_SMALL = 2

' Ranges of valid Timer ID.
Public Const BTR_MIN_TIMERID = 1
Public Const BTR_MAX_TIMERID = 256

' Time unit.
Public Const BTR_TIMEUNIT_MS = 0
Public Const BTR_TIMEUNIT_SC = 1
Public Const BTR_TIMEUNIT_MN = 2
Public Const BTR_TIMEUNIT_HR = 3

'
' TIMER32.DLL API Functions declaration.
'
Public Declare Function vbrcTimerGetCopyright Lib "Timer32.dll" _
    (ByVal ChrCase As Long) As String
Public Declare Function vbrcTimerGetVersion Lib "Timer32.dll" _
    (vMajor As Long, vMinor As Long, vRevision As Long) As String

Public Declare Function vbrcAllocTimer Lib "Timer32.dll" _
    (ByVal TimerID As Long) As Long
Public Declare Function vbrcGetBenchmarkTime Lib "Timer32.dll" _
    (ByVal TimerID As Long, ByVal tmUnit As Long) As Double
Public Declare Function vbrcGetFreeTimerID Lib "Timer32.dll" _
    () As Long
Public Declare Function vbrcGetLoopCounter Lib "Timer32.dll" _
    (ByVal TimerID As Long) As Long
Public Declare Function vbrcGetRemainLoopCount _
    Lib "Timer32.dll" (ByVal TimerID As Long) As Long
Public Declare Function vbrcIsEOLoop Lib "Timer32.dll" _
    (ByVal TimerID As Long) As Boolean
Public Declare Function vbrcReleaseTimer Lib "Timer32.dll" _
    (ByVal TimerID As Long) As Long

Public Declare Sub vbrcDelayTime Lib "Timer32.dll" _
    (ByVal msTime As Long)
Public Declare Sub vbrcInitTimer Lib "Timer32.dll" ()
Public Declare Sub vbrcPauseTimer Lib "Timer32.dll" _
    (ByVal TimerID As Long)
Public Declare Sub vbrcResetTimer Lib "Timer32.dll" _
    (ByVal TimerID As Long)
Public Declare Sub vbrcResumeTimer Lib "Timer32.dll" _
    (ByVal TimerID As Long)
Public Declare Sub vbrcSetLoopCounter Lib "Timer32.dll" _
    (ByVal TimerID As Long, ByVal LoopCnt As Long)
Public Declare Sub vbrcStartTimer Lib "Timer32.dll" _
    (ByVal TimerID As Long)
Public Declare Sub vbrcStopTimer Lib "Timer32.dll" _
    (ByVal TimerID As Long)


