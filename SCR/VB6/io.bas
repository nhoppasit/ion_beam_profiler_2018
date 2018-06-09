Attribute VB_Name = "IO_dll"
Option Explicit

  Public Declare Sub Out Lib "IO.DLL" _
   Alias "PortOut" (ByVal Port As Integer, ByVal Data As Byte)
  
  Public Declare Sub PortWordOut Lib "IO.DLL" _
   (ByVal Port As Integer, ByVal Data As Integer)
  
  Public Declare Sub PortDWordOut Lib "IO.DLL" _
   (ByVal Port As Integer, ByVal Data As Long)
   
  Public Declare Function Inp Lib "IO.DLL" _
   Alias "PortIn" (ByVal Port As Integer) As Byte
   
  Public Declare Function PortWordIn Lib "IO.DLL" _
   (ByVal Port As Integer) As Integer
   
  Public Declare Function PortDWordIn Lib "IO.DLL" _
   (ByVal Port As Integer) As Long
   
  Public Declare Sub SetPortBit Lib "IO.DLL" _
   (ByVal Port As Integer, ByVal Bit As Byte)
   
  Public Declare Sub ClrPortBit Lib "IO.DLL" _
  (ByVal Port As Integer, ByVal Bit As Byte)
  
  Public Declare Sub NotPortBit Lib "IO.DLL" _
   (ByVal Port As Integer, ByVal Bit As Byte)
   
  Public Declare Function GetPortBit Lib "IO.DLL" _
   (ByVal Port As Integer, ByVal Bit As Byte) As Boolean
   
  Public Declare Function RightPortShift Lib "IO.DLL" _
   (ByVal Port As Integer, ByVal Val As Boolean) As Boolean
   
  Public Declare Function LeftPortShift Lib "IO.DLL" _
   (ByVal Port As Integer, ByVal Val As Boolean) As Boolean
  
  Public Declare Function IsDriverInstalled Lib "IO.DLL" () As Boolean

