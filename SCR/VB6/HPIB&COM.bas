Attribute VB_Name = "HPIB_COM"
Option Explicit

Public Const MMC_2_XY_AD = "COM1"              ' MMC-2 Stepping motor driver connected at COM1
Public Const MMC_2_Z_AD = "COM2"              ' MMC-2 Stepping motor driver connected at COM2
Public Const DMM_AD = "gpib0,22"

Public MMC_2XY_ID As Integer
Public MMC_2Z_ID As Integer
Public DMM_ID As Integer


Public Function Read_Instrument(id As Integer) As String
                  
   Dim readbuf As String * 50
   Dim nargs As Integer
   Dim read_cnt As Integer
   
   On Error GoTo ErrorHandler
   
   nargs = ivscanf(id, "%200t", readbuf)
   read_cnt = InStr(1, readbuf, vbLf) - 1
   If Mid$(readbuf, read_cnt, 1) = vbCr Then read_cnt = read_cnt - 1
   Read_Instrument = Left(readbuf, read_cnt)

   Exit Function

ErrorHandler:
'  Close the device session if iopen was successful.
   If id <> 0 Then
      iclose (id)
   End If
   MsgBox Error$, vbOKOnly, "ERROR"
   Read_Instrument = "0"
   Exit Function
End Function
Public Sub Write_Instrument(id As Integer, dev_cmd As String, Optional Terminator As String)
         
   Dim nargs As Integer
   
   nargs = ivprintf(id, dev_cmd & Terminator, 0&)
 
End Sub

Public Function Initialize_Instrument(id As Integer, Instrument_Address As String) As Boolean
                  
   Dim nargs As Integer
   On Error GoTo ErrorHandler

   id = iopen(Instrument_Address)

   Call itimeout(id, 2000)
    
   Initialize_Instrument = True
   
   Exit Function

ErrorHandler:
'  Close the device session if iopen was successful.
   If id <> 0 Then
      iclose (id)
   End If
   MsgBox Error$, vbOKOnly, "Can't Open Instrument"
   Initialize_Instrument = False
   'Exit Function
End Function

Public Sub Finalized_Instrument(id As Integer)
    Call iclose(id)

    '  Tell SICL to cleanup for this task
    'Call siclcleanup
End Sub
