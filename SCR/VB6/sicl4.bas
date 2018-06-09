Attribute VB_Name = "HPIB"
' Copyright 1992-1994 Hewlett-Packard Company.  All Rights Reserved.
'
' This file defines constants, record types, and entry points
' for the HP Standard Instrument Control Library.  You need to
' add this file to each Visual BASIC project that uses the
' HP Standard Instrument Control Library.

' Name of SICL DLL

#If Win16 Then
Const conSiclDll$ = "SICL16.DLL"
#Else
Const conSiclDll$ = "SICL32.DLL"
#End If

' Support levels:
Global Const I_SICL_REVISION = 39       ' HP SICL Revision 3.9
Global Const I_SICL_LEVEL = 3           ' Support Level

' Byte Ordering constants
Global Const I_ORDER_LE = True
Global Const I_ORDER_BE = False

' Session types
Global Const I_SESS_INTF = 1
Global Const I_SESS_DEV = 2
Global Const I_SESS_CMDR = 3

' Interface Types
Global Const I_INTF_NONE = 0
Global Const I_INTF_GPIB = 1
Global Const I_INTF_VXI = 2
Global Const I_INTF_RS232 = 3
Global Const I_INTF_GPIO = 4
' 5 is reserved -- don't use
Global Const I_INTF_USRDEF = 6
' 7 is reserved -- don't use
Global Const I_INTF_MSIB = 8
Global Const I_INTF_LAN = 9

' iread termination conditions
Global Const I_TERM_MAXCNT = 1
Global Const I_TERM_CHR = 2
Global Const I_TERM_END = 4
Global Const I_TERM_NON_BLOCKED = 8

' ixtrig which values.
Global Const I_TRIG_STD = &H1&
Global Const I_TRIG_ALL = &HFFFFFFFF
Global Const I_TRIG_TTL0 = &H1000&
Global Const I_TRIG_TTL1 = &H2000&
Global Const I_TRIG_TTL2 = &H4000&
Global Const I_TRIG_TTL3 = &H8000&
Global Const I_TRIG_TTL4 = &H10000
Global Const I_TRIG_TTL5 = &H20000
Global Const I_TRIG_TTL6 = &H40000
Global Const I_TRIG_TTL7 = &H80000
Global Const I_TRIG_ECL0 = &H100000
Global Const I_TRIG_ECL1 = &H200000
Global Const I_TRIG_ECL2 = &H400000
Global Const I_TRIG_ECL3 = &H800000
Global Const I_TRIG_EXT0 = &H1000000
Global Const I_TRIG_EXT1 = &H2000000
Global Const I_TRIG_EXT2 = &H4000000
Global Const I_TRIG_EXT3 = &H8000000
Global Const I_TRIG_CLK0 = &H10000000
Global Const I_TRIG_CLK1 = &H20000000
Global Const I_TRIG_CLK2 = &H40000000
Global Const I_TRIG_CLK10 = &H80000000
Global Const I_TRIG_CLK100 = &H800&
Global Const I_TRIG_SERIAL_DTR = &H400&
Global Const I_TRIG_SERIAL_RTS = &H200&
Global Const I_TRIG_GPIO_CTL0 = &H100&
Global Const I_TRIG_GPIO_CTL1 = &H80&

' ihint values
Global Const I_HINT_DONTCARE = 0
Global Const I_HINT_USEDMA = 1
Global Const I_HINT_USEPOLL = 2
Global Const I_HINT_USEINTR = 3
Global Const I_HINT_SYSTEM = 4
Global Const I_HINT_IO = 5

' isetintr values.  1-15 are interface independant.
Global Const I_INTR_OFF = 0
Global Const I_INTR_INTFACT = 1
Global Const I_INTR_INTFDEACT = 2
Global Const I_INTR_TRIG = 3
Global Const I_INTR_STB = 4
Global Const I_INTR_DEVCLR = 5

' VXI Interrupts
Global Const I_INTR_VXI_SIGNAL = 16
Global Const I_INTR_VXI_SYSRESET = 17
Global Const I_INTR_VXI_VME = 18
Global Const I_INTR_VXI_LLOCK = 19
Global Const I_INTR_VXI_UKNSIG = 20
Global Const I_INTR_VXI_VMESYSFAIL = 21
Global Const I_INTR_VME_IRQ1 = 22
Global Const I_INTR_VME_IRQ2 = 23
Global Const I_INTR_VME_IRQ3 = 24
Global Const I_INTR_VME_IRQ4 = 25
Global Const I_INTR_VME_IRQ5 = 26
Global Const I_INTR_VME_IRQ6 = 27
Global Const I_INTR_VME_IRQ7 = 28
Global Const I_INTR_ANY_SIG = 29

' GP-IB Interrupts
Global Const I_INTR_GPIB_IFC = 16
Global Const I_INTR_GPIB_PPOLLCONFIG = 17
Global Const I_INTR_GPIB_REMLOC = 18
Global Const I_INTR_GPIB_GET = 20
Global Const I_INTR_GPIB_TLAC = 21

' RS-232 Interrupts
Global Const I_INTR_SERIAL_DAV = 16
Global Const I_INTR_SERIAL_MSL = 17
Global Const I_INTR_SERIAL_BREAK = 18
Global Const I_INTR_SERIAL_ERROR = 19
Global Const I_INTR_SERIAL_TEMT = 20
Global Const I_INTR_SERIAL_MCL = 21

' GP-IO Interrupts
Global Const I_INTR_GPIO_EIR = 16
Global Const I_INTR_GPIO_RDY = 17

' MSIB Interrupts
Global Const I_INTR_MSIB_END_RECEIVED = 22
Global Const I_INTR_MSIB_LINK_BROKEN = 23

' 32 maximum isetintr values
Global Const I_INTR_MAX = 32

' ivxibusstatus values
Global Const I_VXI_BUS_TRIGGER = 0
Global Const I_VXI_BUS_LADDR = 1
Global Const I_VXI_BUS_SERVANT_AREA = 2
Global Const I_VXI_BUS_NORMOP = 3
Global Const I_VXI_BUS_CMDR_LADDR = 4
Global Const I_VXI_BUS_MAN_ID = 5
Global Const I_VXI_BUS_MODEL_ID = 6
Global Const I_VXI_BUS_PROTOCOL = 7
Global Const I_VXI_BUS_XPROT = 8
Global Const I_VXI_BUS_SHM_SIZE = 9
Global Const I_VXI_BUS_SHM_ADDR_SPACE = 10
Global Const I_VXI_SHM_PAGE = 11
Global Const I_VXI_BUS_VXIMXI = 12
Global Const I_VXI_BUS_TRIGSUPP = 13

' igpibbusstatus values
Global Const I_GPIB_BUS_REM = 1
Global Const I_GPIB_BUS_SRQ = 2
Global Const I_GPIB_BUS_NDAC = 3
Global Const I_GPIB_BUS_SYSCTLR = 4
Global Const I_GPIB_BUS_ACTCTLR = 5
Global Const I_GPIB_BUS_TALKER = 6
Global Const I_GPIB_BUS_LISTENER = 7
Global Const I_GPIB_BUS_ADDR = 8
Global Const I_GPIB_BUS_LINES = 9

Global Const I_GPIB_T1DELAY_MIN = 350
Global Const I_GPIB_T1DELAY_MAX = 2400
   
' values for igpioctrl and igpiostat
Global Const I_GPIO_AUX = 1
Global Const I_GPIO_CTRL = 2
Global Const I_GPIO_DATA = 3
Global Const I_GPIO_INFO = 4
Global Const I_GPIO_SET_PCTL = 5
Global Const I_GPIO_STAT = 6
Global Const I_GPIO_READ_EOI = 7
Global Const I_GPIO_TEST_ONLY = 8
Global Const I_GPIO_POLARITY = 9
Global Const I_GPIO_READ_CLK = 10
Global Const I_GPIO_PCTL_DELAY = 11

Global Const I_GPIO_CTRL_CTL0 = &H1
Global Const I_GPIO_CTRL_CTL1 = &H2

Global Const I_GPIO_STAT_STI0 = &H1
Global Const I_GPIO_STAT_STI1 = &H2
Global Const I_GPIO_EIR = &H4
Global Const I_GPIO_PSTS = &H8
Global Const I_GPIO_CHK_PSTS = &H10
Global Const I_GPIO_AUTO_HDSK = &H20
Global Const I_GPIO_ENH_MODE = &H40
Global Const I_GPIO_READY = &H80
Global Const I_GPIO_EOI_NONE = &H10000

' RS-232 values
Global Const I_SERIAL_BAUD = 1
Global Const I_SERIAL_PARITY = 2
Global Const I_SERIAL_STOP = 3
Global Const I_SERIAL_WIDTH = 4
Global Const I_SERIAL_FLOW_CTRL = 5
Global Const I_SERIAL_MSL = 6
Global Const I_SERIAL_STAT = 7
Global Const I_SERIAL_RESET = 9
Global Const I_SERIAL_READ_EOI = 10
Global Const I_SERIAL_WRITE_EOI = 11
Global Const I_SERIAL_DUPLEX = 12
Global Const I_SERIAL_READ_BUFSZ = 13
Global Const I_SERIAL_READ_DAV = 14

' RS-232 duplex modes
Global Const I_SERIAL_DUPLEX_HALF = 1
Global Const I_SERIAL_DUPLEX_FULL = 2

' RS-232 UART status
Global Const I_SERIAL_DAV = &H1
Global Const I_SERIAL_OVERFLOW = &H2
Global Const I_SERIAL_PARERR = &H4
Global Const I_SERIAL_FRAMING = &H8
Global Const I_SERIAL_BREAK = &H10
Global Const I_SERIAL_TEMT = &H20

' RS-232 flow control
Global Const I_SERIAL_FLOW_NONE = 0
Global Const I_SERIAL_FLOW_XON = 1
Global Const I_SERIAL_FLOW_RTS_CTS = 2
Global Const I_SERIAL_FLOW_DTR_DSR = 3

' RS-232 modem status lines
Global Const I_SERIAL_DCD = &H1
Global Const I_SERIAL_DSR = &H2
Global Const I_SERIAL_CTS = &H4
Global Const I_SERIAL_RI = &H8
Global Const I_SERIAL_D_DCD = &H10
Global Const I_SERIAL_D_DSR = &H20
Global Const I_SERIAL_D_CTS = &H40
Global Const I_SERIAL_D_TERI = &H80

' RS-232 modem control lines
Global Const I_SERIAL_RTS = &H1000
Global Const I_SERIAL_DTR = &H2000

' RS-232 parity values
Global Const I_SERIAL_PAR_NONE = 0
Global Const I_SERIAL_PAR_EVEN = 1
Global Const I_SERIAL_PAR_ODD = 2
Global Const I_SERIAL_PAR_MARK = 3
Global Const I_SERIAL_PAR_SPACE = 4
Global Const I_SERIAL_PAR_IGNORE = 5

' RS-232 stop-bit values
Global Const I_SERIAL_STOP_1 = 1
Global Const I_SERIAL_STOP_2 = 2

' RS-232 character width
Global Const I_SERIAL_CHAR_5 = 5
Global Const I_SERIAL_CHAR_6 = 6
Global Const I_SERIAL_CHAR_7 = 7
Global Const I_SERIAL_CHAR_8 = 8

' EOI support (used with the I_SERIAL_*_EOI command)
Global Const I_SERIAL_EOI_CHR = &H100
Global Const I_SERIAL_EOI_NONE = &H200
Global Const I_SERIAL_EOI_BIT8 = &H400


' MSIB error types (for imsibseterror)
Global Const I_MSIB_PERMANENTERR = 0
Global Const I_MSIB_TRANSIENTERR = 1

' MSIB commands (for imsibcmd)
Global Const I_MSIB_CMD_NULL = &H0
Global Const I_MSIB_CMD_END = &H1
Global Const I_MSIB_CMD_SEND_CAPABILITY = &H2
Global Const I_MSIB_CMD_RETURN_TO_LOCAL = &H6
Global Const I_MSIB_CMD_LOCK_LINK = &H7
Global Const I_MSIB_CMD_UNLOCK_LINK = &H8
Global Const I_MSIB_CMD_LIGHT_ACTIVE = &H9
Global Const I_MSIB_CMD_UNLIGHT_ACTIVE = &HA
Global Const I_MSIB_CMD_ERROR_OCCURRED = &HB
Global Const I_MSIB_CMD_ERRORS_CLEARED = &HC
Global Const I_MSIB_CMD_SEND_STATUS = &H10
Global Const I_MSIB_CMD_SEND_ERRORS = &H11
Global Const I_MSIB_CMD_SEND_MODULE_ID = &H12
Global Const I_MSIB_CMD_SEND_MANUFACTURER = &H13
Global Const I_MSIB_CMD_SEND_TIME = &H14
Global Const I_MSIB_CMD_LINK_REMOTE = &H15
Global Const I_MSIB_CMD_LINK_LOCAL = &H16
Global Const I_MSIB_CMD_SEND_MODEL_NUMBER = &H17
Global Const I_MSIB_CMD_SEND_SERIAL_NUMBER = &H18
Global Const I_MSIB_CMD_SEND_FIRMWARE_REV = &H19
Global Const I_MSIB_CMD_STATUS = &H600
Global Const I_MSIB_CMD_SET_IEEE_ADDRESS = &H700

' imap mapspace values
Global Const I_MAP_A16 = &H0
Global Const I_MAP_A24 = &H1
Global Const I_MAP_A32 = &H2
Global Const I_MAP_VXIDEV = &H3
Global Const I_MAP_EXTEND = &H4
Global Const I_MAP_INTFREG = &H5
Global Const I_MAP_SHARED = &H6

' Following is for icmd; uses Radisys define
Global Const DOCMD_VALIDATE_MAPPING = &H40000005

' Error Codes
' NOTE that User Error Codes 32501-32630 are reserved
' for HP SICL.
Const SICL_ERR_BASE = 32501

Global Const I_ERR_NOERROR = 0
Global Const I_ERR_SYNTAX = SICL_ERR_BASE
Global Const I_ERR_SYMNAME = 1 + SICL_ERR_BASE
Global Const I_ERR_BADADDR = 2 + SICL_ERR_BASE
Global Const I_ERR_BADID = 3 + SICL_ERR_BASE
Global Const I_ERR_PARAM = 4 + SICL_ERR_BASE
Global Const I_ERR_NOCONN = 5 + SICL_ERR_BASE
Global Const I_ERR_NOPERM = 6 + SICL_ERR_BASE
Global Const I_ERR_NOTSUPP = 7 + SICL_ERR_BASE
Global Const I_ERR_NORSRC = 8 + SICL_ERR_BASE
Global Const I_ERR_NOINTF = 9 + SICL_ERR_BASE
Global Const I_ERR_LOCKED = 10 + SICL_ERR_BASE
Global Const I_ERR_NOLOCK = 11 + SICL_ERR_BASE
Global Const I_ERR_BADFMT = 12 + SICL_ERR_BASE
Global Const I_ERR_DATA = 13 + SICL_ERR_BASE
Global Const I_ERR_TIMEOUT = 14 + SICL_ERR_BASE
Global Const I_ERR_OVERFLOW = 15 + SICL_ERR_BASE
Global Const I_ERR_IO = 16 + SICL_ERR_BASE
Global Const I_ERR_OS = 17 + SICL_ERR_BASE
Global Const I_ERR_BADMAP = 18 + SICL_ERR_BASE
Global Const I_ERR_NODEV = 19 + SICL_ERR_BASE
Global Const I_ERR_INVLADDR = 20 + SICL_ERR_BASE
Global Const I_ERR_NOTIMPL = 21 + SICL_ERR_BASE
Global Const I_ERR_ABORTED = 22 + SICL_ERR_BASE
Global Const I_ERR_BADCONFIG = 23 + SICL_ERR_BASE
Global Const I_ERR_NOCMDR = 24 + SICL_ERR_BASE
Global Const I_ERR_VERSION = 25 + SICL_ERR_BASE
Global Const I_ERR_NESTEDIO = 26 + SICL_ERR_BASE
Global Const I_ERR_BUSY = 27 + SICL_ERR_BASE
Global Const I_ERR_CONNEXISTS = 28 + SICL_ERR_BASE
Global Const I_ERR_BUSERR = 29 + SICL_ERR_BASE
Global Const I_ERR_BUSERR_RETRY = 30 + SICL_ERR_BASE
Global Const I_ERR_INTERNAL = 127 + SICL_ERR_BASE
Global Const I_ERR_INTERRUPT = 128 + SICL_ERR_BASE
Global Const I_ERR_UNKNOWNERR = 129 + SICL_ERR_BASE
Global Const SICL_ERR_LAST = I_ERR_UNKNOWNERR

Global Const I_READ_BUF_SZ = 4096
Global Const I_WRITE_BUF_SZ = 128

Global Const I_BUF_READ = &H1
Global Const I_BUF_WRITE = &H2
Global Const I_BUF_DISCARD_READ = &H4
Global Const I_BUF_DISCARD_WRITE = &H8
Global Const I_BUF_WRITE_END = &H10

' Data Types used by SICL
Type lu_info
  logical_unit As Long
  symname As String * 32
  cardname As String * 32
  filler As Long
  intftype As Long
  location As Long
  busaddr As Long
  hwarg(0 To 15) As String * 20
  visaname As String * 32
  filler2(0 To 3) As Long
End Type

Type vxiinfo
  laddr As Integer
  name As String * 16
  manuf_name As String * 16
  model_name As String * 16
  man_id As Long
  model As Long
  devclass As Long
  selftest As Integer
  cage_num As Integer
  slot As Integer
  protocol As Long
  x_protocol As Long
  servant_area As Long
  addrspace As Long
  memsize As Long
  memstart As Long
  slot0_laddr As Integer
  cmdr_laddr As Integer
  int_handler(0 To 7)  As Integer
  interrupter(0 To 7) As Integer
  fill(0 To 9) As Integer
End Type

#If Win16 Then
    
' Version Information
Declare Function vb_iversion Lib "vbsicl16.dll" (specversion As Integer, implversion As Integer) As Integer
Declare Function vb_idrvrversion Lib "vbsicl16.dll" (ByVal id As Integer, specversion As Integer, implversion As Integer) As Integer

' Open/Close
Declare Function vb_iopen Lib "vbsicl16.dll" (ByVal addr As String) As Integer
Declare Function vb_iclose Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_igetintfsess Lib "vbsicl16.dll" (ByVal id As Integer) As Integer

' Write/Read

Declare Function vb_iwrite Lib "vbsicl16.dll" (ByVal which As Integer, ByVal id As Integer, ByVal buf As Variant, ByVal datalen As Long, ByVal endi As Integer, actual As Long) As Integer
Declare Function vb_iread Lib "vbsicl16.dll" (ByVal which As Integer, ByVal id As Integer, buf As Variant, ByVal bufsize As Long, reason As Integer, actual As Long) As Integer
Declare Function vb_itermchr Lib "vbsicl16.dll" (ByVal id As Integer, ByVal tchr As Integer) As Integer
Declare Function vb_igettermchr Lib "vbsicl16.dll" (ByVal id As Integer, tchr As Integer) As Integer

' Formatted I/O
Declare Function vb_ivprintf Lib "vbsicl16.dll" (ByVal id As Integer, ByVal fmt As String, ByVal ap As Variant, ByVal lenBstr As Long) As Integer
Declare Function vb_ivscanf Lib "vbsicl16.dll" (ByVal id As Integer, ByVal fmt As String, ByRef ap As Variant, ByVal lenBstr As Long) As Integer
Declare Function vb_iflush Lib "vbsicl16.dll" (ByVal id As Integer, ByVal mask As Integer) As Integer
Declare Function vb_isetbuf Lib "vbsicl16.dll" (ByVal id As Integer, ByVal mask As Integer, ByVal size As Integer) As Integer

' Device/Interface Control
Declare Function vb_iclear Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_ilocal Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_iremote Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_ireadstb Lib "vbsicl16.dll" (ByVal id As Integer, ByRef stb As Integer) As Integer
Declare Function vb_itrigger Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_ixtrig Lib "vbsicl16.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ihint Lib "vbsicl16.dll" (ByVal id As Integer, ByVal hint As Integer) As Integer

' Commander Sessions
Declare Function vb_isetstb Lib "vbsicl16.dll" (ByVal id As Integer, ByVal stb As Byte) As Integer

' Locking
Declare Function vb_ilock Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_iunlock Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_isetlockwait Lib "vbsicl16.dll" (ByVal id As Integer, ByVal flag As Integer) As Integer
Declare Function vb_igetlockwait Lib "vbsicl16.dll" (ByVal id As Integer, flag As Integer) As Integer

' Timeouts
Declare Function vb_itimeout Lib "vbsicl16.dll" (ByVal id As Integer, ByVal tval As Long) As Integer
Declare Function vb_igettimeout Lib "vbsicl16.dll" (ByVal id As Integer, tval As Long) As Integer

' Misc routines
Declare Function vb_igetaddr Lib "vbsicl16.dll" (ByVal id As Integer, ByVal addr As String) As Integer
Declare Function vb_igetintftype Lib "vbsicl16.dll" (ByVal id As Integer, pdata As Integer) As Integer
Declare Function vb_igetsesstype Lib "vbsicl16.dll" (ByVal id As Integer, pdata As Integer) As Integer
Declare Function vb_igetdevaddr Lib "vbsicl16.dll" (ByVal id As Integer, prim As Integer, sec As Integer) As Integer
Declare Function vb_igetlu Lib "vbsicl16.dll" (ByVal id As Integer, lu As Integer) As Integer
Declare Function vb_iswap Lib "vbsicl16.dll" (ByRef addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
Declare Function vb_igetlulist Lib "vbsicl16.dll" (list() As Integer) As Integer
Declare Function vb_igetluinfo Lib "vbsicl16.dll" (ByVal lu As Integer, result As lu_info) As Integer
Declare Function vb_igetgatewaytype Lib "vbsicl16.dll" (ByVal id As Integer, pdata As Integer) As Integer

' Error Handling
Declare Function vb_igeterrno Lib "vbsicl16.dll" () As Integer
Declare Function vb_iseterrno Lib "vbsicl16.dll" (ByVal id As Integer, ByVal Errno As Integer) As Integer
Declare Function vb_igeterrstr Lib "vbsicl16.dll" (ByVal errcode As Integer, ByVal myerrstr As String) As Integer
Declare Function vb_icauseerr Lib "vbsicl16.dll" (ByVal id As Integer, ByVal errcode As Integer, ByVal flag As Integer) As Integer
Declare Function vbsetsiclerrbase Lib "vbsicl16.dll" (ByVal errbase As Integer) As Integer

' RS-232 specific routines
Declare Function vb_iserialmclctrl Lib "vbsicl16.dll" (ByVal id As Integer, ByVal sline As Integer, ByVal state As Integer) As Integer
Declare Function vb_iserialmclstat Lib "vbsicl16.dll" (ByVal id As Integer, ByVal sline As Integer, state As Integer) As Integer
Declare Function vb_iserialctrl Lib "vbsicl16.dll" (ByVal id As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
Declare Function vb_iserialstat Lib "vbsicl16.dll" (ByVal id As Integer, ByVal request As Integer, result As Long) As Integer
Declare Function vb_iserialbreak Lib "vbsicl16.dll" (ByVal id As Integer) As Integer

' VXI Specific routines
Declare Function vb_ivxibusstatus Lib "vbsicl16.dll" (ByVal id As Integer, ByVal request As Integer, result As Long) As Integer
Declare Function vb_ivxiwaitnormop Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_ivxitrigon Lib "vbsicl16.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ivxitrigoff Lib "vbsicl16.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ivxitrigroute Lib "vbsicl16.dll" (ByVal id As Integer, ByVal in_which As Long, ByVal out_which As Long) As Integer
Declare Function vb_ivxigettrigroute Lib "vbsicl16.dll" (ByVal id As Integer, ByVal which As Long, route As Long) As Integer
Declare Function vb_ivxiws Lib "vbsicl16.dll" (ByVal id As Integer, ByVal wscmd As Integer, wsresp As Integer, rpe As Integer) As Integer
Declare Function vb_ivxiservants Lib "vbsicl16.dll" (ByVal id As Integer, ByVal maxnum As Integer, list() As Integer) As Integer
Declare Function vb_ivxirminfo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal laddr As Integer, ByRef info As vxiinfo) As Integer

' GP-IB Specific Details
Declare Function vb_igpibbusstatus Lib "vbsicl16.dll" (ByVal id As Integer, ByVal request As Integer, result As Integer) As Integer
Declare Function vb_igpibppoll Lib "vbsicl16.dll" (ByVal id As Integer, result As Integer) As Integer
Declare Function vb_igpibppollconfig Lib "vbsicl16.dll" (ByVal id As Integer, ByVal cval As Integer) As Integer
Declare Function vb_igpibppollresp Lib "vbsicl16.dll" (ByVal id As Integer, ByVal sval As Integer) As Integer
Declare Function vb_igpibpassctl Lib "vbsicl16.dll" (ByVal id As Integer, ByVal busaddr As Integer) As Integer
Declare Function vb_igpibrenctl Lib "vbsicl16.dll" (ByVal id As Integer, ByVal ren As Integer) As Integer
Declare Function vb_igpibatnctl Lib "vbsicl16.dll" (ByVal id As Integer, ByVal atnval As Integer) As Integer
Declare Function vb_igpibsendcmd Lib "vbsicl16.dll" (ByVal id As Integer, ByVal buf As String, ByVal length As Integer) As Integer
Declare Function vb_igpibllo Lib "vbsicl16.dll" (ByVal id As Integer) As Integer
Declare Function vb_igpibbusaddr Lib "vbsicl16.dll" (ByVal id As Integer, ByVal busaddr As Integer) As Integer
Declare Function vb_igpibgett1delay Lib "vbsicl16.dll" (ByVal id As Integer, delay As Integer) As Integer
Declare Function vb_igpibsett1delay Lib "vbsicl16.dll" (ByVal id As Integer, ByVal delay As Integer) As Integer
Declare Function vb_igpibpulseifc Lib "vbsicl16.dll" (ByVal id As Integer) As Integer

' GPIO Specific routines
Declare Function vb_igpioctrl Lib "vbsicl16.dll" (ByVal id As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
Declare Function vb_igpiostat Lib "vbsicl16.dll" (ByVal id As Integer, ByVal request As Integer, ByRef result As Long) As Integer
Declare Function vb_igpiosetwidth Lib "vbsicl16.dll" (ByVal id As Integer, ByVal dwidth As Integer) As Integer
Declare Function vb_igpiogetwidth Lib "vbsicl16.dll" (ByVal id As Integer, ByRef dwidth As Integer) As Integer

' LAN Specific functions
Declare Function vb_ilantimeout Lib "vbsicl16.dll" (ByVal id As Integer, ByVal tval As Long) As Integer
Declare Function vb_ilangettimeout Lib "vbsicl16.dll" (ByVal id As Integer, tval As Long) As Integer

' Map routines
Declare Function vb_imap Lib "vbsicl16.dll" (ByVal id As Integer, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer, ByVal suggested As Long) As Long
Declare Function vb_iunmap Lib "vbsicl16.dll" (ByVal id As Integer, ByVal addr As Long, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
Declare Function vb_imapinfo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal mapspace As Integer, numwindows As Integer, winsize As Integer) As Integer

' Block copy and fifo routines
Declare Function vb_ibblockcopy Lib "vbsicl16.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long) As Integer
Declare Function vb_iwblockcopy Lib "vbsicl16.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilblockcopy Lib "vbsicl16.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ibpushfifo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long) As Integer
Declare Function vb_iwpushfifo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilpushfifo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ibpopfifo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long) As Integer
Declare Function vb_iwpopfifo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilpopfifo Lib "vbsicl16.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_icmd Lib "vbsicl16.dll" (ByVal id As Integer, ByVal cmd As Long, ByVal datalen As Integer, ByVal datawidth As Integer, ByRef pdata As Long) As Integer

' Windows 3.1 Cleanup routines
Declare Function vb__siclcleanup Lib "vbsicl16.dll" () As Integer

' Windows 3.1 yield control routine
Declare Function vb__setsiclyield Lib "vbsicl16.dll" (ByVal yield_option As Integer) As Integer

' Peek/Poke routines
Declare Sub vb_ibpoke Lib "vbsicl16.dll" (ByVal addr As Long, ByVal Value As Byte)
Declare Sub vb_iwpoke Lib "vbsicl16.dll" (ByVal addr As Long, ByVal Value As Integer)
Declare Sub vb_ilpoke Lib "vbsicl16.dll" (ByVal addr As Long, ByVal Value As Long)
Declare Function vb_ibpeek Lib "vbsicl16.dll" (ByVal addr As Long) As Byte
Declare Function vb_iwpeek Lib "vbsicl16.dll" (ByVal addr As Long) As Integer
Declare Function vb_ilpeek Lib "vbsicl16.dll" (ByVal addr As Long) As Long

#Else

' Version Information

Declare Function vb_iversion Lib "vbsicl32.dll" (specversion As Integer, implversion As Integer) As Integer
Declare Function vb_idrvrversion Lib "vbsicl32.dll" (ByVal id As Integer, specversion As Integer, implversion As Integer) As Integer

' Open/Close
Declare Function vb_iopen Lib "vbsicl32.dll" (ByVal addr As String) As Integer
Declare Function vb_iclose Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_igetintfsess Lib "vbsicl32.dll" (ByVal id As Integer) As Integer

' Write/Read

Declare Function vb_iwrite Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, ByVal buf As Variant, ByVal datalen As Long, ByVal endi As Integer, actual As Long) As Integer
Declare Function vb_iread Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, buf As Variant, ByVal bufsize As Long, reason As Integer, actual As Long) As Integer
Declare Function vb_itermchr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal tchr As Integer) As Integer
Declare Function vb_igettermchr Lib "vbsicl32.dll" (ByVal id As Integer, tchr As Integer) As Integer

' Formatted I/O
Declare Function vb_iscan Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, ByVal s As String, ByVal fmt As String, ByRef va1 As Variant, ByRef va2 As Variant, ByRef va3 As Variant, ByRef va4 As Variant, ByRef va5 As Variant, ByRef va6 As Variant, ByRef va7 As Variant, ByRef va8 As Variant, ByRef va9 As Variant, ByRef va10 As Variant) As Integer
Declare Function vb_iprint Lib "vbsicl32.dll" (ByVal which As Integer, ByVal id As Integer, ByVal s As String, ByVal fmt As String, ByRef ap() As Variant) As Integer

Declare Function vb_ivprintf Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fmt As String, ByVal ap As Variant, ByVal lenBstr As Long) As Integer
Declare Function vb_ivscanf Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fmt As String, ByRef ap As Variant, ByVal lenBstr As Long) As Integer
Declare Function vb_iflush Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mask As Integer) As Integer
Declare Function vb_isetbuf Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mask As Integer, ByVal size As Integer) As Integer

' Device/Interface Control
Declare Function vb_iclear Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ilocal Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_iremote Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ireadstb Lib "vbsicl32.dll" (ByVal id As Integer, ByRef stb As Integer) As Integer
Declare Function vb_itrigger Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ixtrig Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ihint Lib "vbsicl32.dll" (ByVal id As Integer, ByVal hint As Integer) As Integer

' Commander Sessions
Declare Function vb_isetstb Lib "vbsicl32.dll" (ByVal id As Integer, ByVal stb As Byte) As Integer

' Locking
Declare Function vb_ilock Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_iunlock Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_isetlockwait Lib "vbsicl32.dll" (ByVal id As Integer, ByVal flag As Integer) As Integer
Declare Function vb_igetlockwait Lib "vbsicl32.dll" (ByVal id As Integer, flag As Integer) As Integer

' Timeouts
Declare Function vb_itimeout Lib "vbsicl32.dll" (ByVal id As Integer, ByVal tval As Long) As Integer
Declare Function vb_igettimeout Lib "vbsicl32.dll" (ByVal id As Integer, tval As Long) As Integer

' Misc routines
Declare Function vb_igetaddr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal addr As String) As Integer
Declare Function vb_igetintftype Lib "vbsicl32.dll" (ByVal id As Integer, pdata As Integer) As Integer
Declare Function vb_igetsesstype Lib "vbsicl32.dll" (ByVal id As Integer, pdata As Integer) As Integer
Declare Function vb_igetdevaddr Lib "vbsicl32.dll" (ByVal id As Integer, prim As Integer, sec As Integer) As Integer
Declare Function vb_igetlu Lib "vbsicl32.dll" (ByVal id As Integer, lu As Integer) As Integer
Declare Function vb_iswap Lib "vbsicl32.dll" (ByRef addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
Declare Function vb_igetlulist Lib "vbsicl32.dll" (list() As Integer) As Integer
Declare Function vb_igetluinfo Lib "vbsicl32.dll" (ByVal lu As Integer, result As lu_info) As Integer
Declare Function vb_igetgatewaytype Lib "vbsicl32.dll" (ByVal id As Integer, pdata As Integer) As Integer


' Error Handling
Declare Function vb_igeterrno Lib "vbsicl32.dll" () As Integer
Declare Function vb_iseterrno Lib "vbsicl32.dll" (ByVal id As Integer, ByVal xint As Integer) As Integer
Declare Function vb_igeterrstr Lib "vbsicl32.dll" (ByVal errcode As Integer, ByVal myerrstr As String) As Integer
Declare Function vb_icauseerr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal errcode As Integer, ByVal flag As Integer) As Integer
Declare Function vbsetsiclerrbase Lib "vbsicl32.dll" (ByVal errbase As Integer) As Integer

' RS-232 specific routines
Declare Function vb_iserialmclctrl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal sline As Integer, ByVal state As Integer) As Integer
Declare Function vb_iserialmclstat Lib "vbsicl32.dll" (ByVal id As Integer, ByVal sline As Integer, state As Integer) As Integer
Declare Function vb_iserialctrl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
Declare Function vb_iserialstat Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, result As Long) As Integer
Declare Function vb_iserialbreak Lib "vbsicl32.dll" (ByVal id As Integer) As Integer

' VXI Specific routines
Declare Function vb_ivxibusstatus Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, result As Long) As Integer
Declare Function vb_ivxiwaitnormop Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_ivxitrigon Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ivxitrigoff Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long) As Integer
Declare Function vb_ivxitrigroute Lib "vbsicl32.dll" (ByVal id As Integer, ByVal in_which As Long, ByVal out_which As Long) As Integer
Declare Function vb_ivxigettrigroute Lib "vbsicl32.dll" (ByVal id As Integer, ByVal which As Long, route As Long) As Integer
Declare Function vb_ivxiws Lib "vbsicl32.dll" (ByVal id As Integer, ByVal wscmd As Integer, wsresp As Integer, rpe As Integer) As Integer
Declare Function vb_ivxiservants Lib "vbsicl32.dll" (ByVal id As Integer, ByVal maxnum As Integer, list() As Integer) As Integer
Declare Function vb_ivxirminfo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal laddr As Integer, ByRef info As vxiinfo) As Integer

' GP-IB Specific Details
Declare Function vb_igpibbusstatus Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, result As Integer) As Integer
Declare Function vb_igpibppoll Lib "vbsicl32.dll" (ByVal id As Integer, result As Integer) As Integer
Declare Function vb_igpibppollconfig Lib "vbsicl32.dll" (ByVal id As Integer, ByVal cval As Integer) As Integer
Declare Function vb_igpibppollresp Lib "vbsicl32.dll" (ByVal id As Integer, ByVal sval As Integer) As Integer
Declare Function vb_igpibpassctl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal busaddr As Integer) As Integer
Declare Function vb_igpibrenctl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal ren As Integer) As Integer
Declare Function vb_igpibatnctl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal atnval As Integer) As Integer
Declare Function vb_igpibsendcmd Lib "vbsicl32.dll" (ByVal id As Integer, ByVal buf As String, ByVal length As Integer) As Integer
Declare Function vb_igpibllo Lib "vbsicl32.dll" (ByVal id As Integer) As Integer
Declare Function vb_igpibbusaddr Lib "vbsicl32.dll" (ByVal id As Integer, ByVal busaddr As Integer) As Integer
Declare Function vb_igpibgett1delay Lib "vbsicl32.dll" (ByVal id As Integer, delay As Integer) As Integer
Declare Function vb_igpibsett1delay Lib "vbsicl32.dll" (ByVal id As Integer, ByVal delay As Integer) As Integer
Declare Function vb_igpibpulseifc Lib "vbsicl32.dll" (ByVal id As Integer) As Integer

' GPIO Specific routines
Declare Function vb_igpioctrl Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
Declare Function vb_igpiostat Lib "vbsicl32.dll" (ByVal id As Integer, ByVal request As Integer, ByRef result As Long) As Integer
Declare Function vb_igpiosetwidth Lib "vbsicl32.dll" (ByVal id As Integer, ByVal dwidth As Integer) As Integer
Declare Function vb_igpiogetwidth Lib "vbsicl32.dll" (ByVal id As Integer, ByRef dwidth As Integer) As Integer

' LAN Specific functions
Declare Function vb_ilantimeout Lib "vbsicl32.dll" (ByVal id As Integer, ByVal tval As Long) As Integer
Declare Function vb_ilangettimeout Lib "vbsicl32.dll" (ByVal id As Integer, tval As Long) As Integer

' Map routines
Declare Function vb_imap Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer, ByVal suggested As Long) As Long
Declare Function vb_iunmap Lib "vbsicl32.dll" (ByVal id As Integer, ByVal addr As Long, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
Declare Function vb_imapx Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Long
Declare Function vb_iunmapx Lib "vbsicl32.dll" (ByVal id As Integer, ByVal addr As Long, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
Declare Function vb_imapinfo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal mapspace As Integer, numwindows As Integer, winsize As Integer) As Integer

' peekx/pokex/blockmovex routines
Declare Function vb_ipokex8 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal offset As Long, ByVal Value As Byte) As Integer
Declare Function vb_ipokex16 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal offset As Long, ByVal Value As Integer) As Integer
Declare Function vb_ipokex32 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal offset As Long, ByVal Value As Long) As Integer
Declare Function vb_ipeekx8 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal offset As Long, Value As Byte) As Integer
Declare Function vb_ipeekx16 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal offset As Long, Value As Integer) As Integer
Declare Function vb_ipeekx32 Lib "vbsicl32.dll" (ByVal id As Integer, ByVal handle As Long, ByVal offset As Long, Value As Long) As Integer
Declare Function vb_iblockmovex Lib "vbsicl32.dll" (ByVal id As Integer, ByVal srcHandle As Long, ByRef srcOffset As Variant, ByVal srcWidth As Integer, ByVal srcIncrement As Integer, ByVal destHandle As Long, ByRef destOffset As Variant, ByVal destWidth As Integer, ByVal destIncrement As Integer, ByVal count As Long, ByVal swap As Integer) As Integer

' Block copy and fifo routines
Declare Function vb_ibblockcopy Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long) As Integer
Declare Function vb_iwblockcopy Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilblockcopy Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ibpushfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long) As Integer
Declare Function vb_iwpushfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilpushfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ibpopfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long) As Integer
Declare Function vb_iwpopfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_ilpopfifo Lib "vbsicl32.dll" (ByVal id As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
Declare Function vb_icmd Lib "vbsicl32.dll" (ByVal id As Integer, ByVal cmd As Long, ByVal datalen As Integer, ByVal datawidth As Integer, ByRef pdata As Long) As Integer

' Windows 3.1 Cleanup routines
Declare Function vb__siclcleanup Lib "vbsicl32.dll" () As Integer

' Windows 3.1 yield control routine
Declare Function vb__setsiclyield Lib "vbsicl32.dll" (ByVal yield_option As Integer) As Integer

' Peek/Poke routines
Declare Sub vb_ibpoke Lib "vbsicl32.dll" (ByVal addr As Long, ByVal Value As Byte)
Declare Sub vb_iwpoke Lib "vbsicl32.dll" (ByVal addr As Long, ByVal Value As Integer)
Declare Sub vb_ilpoke Lib "vbsicl32.dll" (ByVal addr As Long, ByVal Value As Long)
Declare Function vb_ibpeek Lib "vbsicl32.dll" (ByVal addr As Long) As Byte
Declare Function vb_iwpeek Lib "vbsicl32.dll" (ByVal addr As Long) As Integer
Declare Function vb_ilpeek Lib "vbsicl32.dll" (ByVal addr As Long) As Long
#End If

Function iversion(specversion As Integer, implversion As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iversion(specversion, implversion)
    iversion = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function idrvrversion(id1 As Integer, specversion As Integer, implversion As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_idrvrversion(id1, specversion, implversion)
    idrvrversion = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iopen(siclAddr As String) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iopen(siclAddr)
    iopen = id

    ' If we get 0 back, there was an error, try to report it
    If id = 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iclose(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iclose(id1)
    iclose = id
    
    ' If return value was not 0, we had an error
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetintfsess(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetintfsess(id1)
    igetintfsess = id

    ' If we get 0 back, there was an error, try to report it
    If id = 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwrite(ByVal id1 As Integer, ByVal buf As Variant, ByVal datalen As Long, ByVal endi As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    tmp = VarType(buf)

    'If the buf is a string, then Win16 requires it to be < 32768 bytes
    If tmp = 8 Then
#If Win16 Then
       If datalen > 32767 Then
          Err.Clear
          myerrstr = "Second param string length must be <= 32767"
          Err.Description = myerrstr
          MsgBox myerrstr
          Err.Raise (I_ERR_PARAM) 'Raise the error
       End If
#End If
    End If

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwrite(1, id1, buf, datalen, endi, actual)
    
    iwrite = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ifwrite(ByVal id1 As Integer, ByVal buf As Variant, ByVal datalen As Long, ByVal endi As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    tmp = VarType(buf)

    'If the buf is a string, then Win16 requires it to be < 32768 bytes
    If tmp = 8 Then
#If Win16 Then
       If datalen > 32767 Then
          Err.Clear
          myerrstr = "Second param string length must be <= 32767"
          Err.Description = myerrstr
          MsgBox myerrstr
          Err.Raise (I_ERR_PARAM) 'Raise the error
       End If
#End If
    End If

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwrite(2, id1, buf, datalen, endi, actual)
    
    ifwrite = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function iread(ByVal id1 As Integer, ByRef buf As Variant, ByVal bufsize As Long, reason As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
        
    tmp = VarType(buf)

    'If the buf parameter string, Win16 needs it to be < 32768 bytes
    If tmp = 8 Then
#If Win16 Then
       If bufsize > 32767 Then
          Err.Clear
          myerrstr = "Second param string length must be <= 32767"
          Err.Description = myerrstr
          MsgBox myerrstr
          Err.Raise (I_ERR_PARAM) 'Raise the error
       End If
#End If
    End If

    ' Call the function in the SICL DLL and check for errors
    id = vb_iread(1, id1, buf, bufsize, reason, actual)
    
    iread = id

    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ifread(ByVal id1 As Integer, ByRef buf As Variant, ByVal bufsize As Long, reason As Integer, actual As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
        
    tmp = VarType(buf)

    'If the buf parameter string, Win16 needs it to be < 32768 bytes
    If tmp = 8 Then
#If Win16 Then
       If bufsize > 32767 Then
          Err.Clear
          myerrstr = "Second param string length must be <= 32767"
          Err.Description = myerrstr
          MsgBox myerrstr
          Err.Raise (I_ERR_PARAM) 'Raise the error
       End If
#End If
    End If

    ' Call the function in the SICL DLL and check for errors
    id = vb_iread(2, id1, buf, bufsize, reason, actual)
    
    ifread = id

    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function itermchr(ByVal id1 As Integer, ByVal tchr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_itermchr(id1, tchr)
    itermchr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igettermchr(ByVal id1 As Integer, tchr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igettermchr(id1, tchr)
    igettermchr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function ivprintf(ByVal id1 As Integer, ByVal fmt As String, Optional ByVal ap As Variant) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    Dim howLong As Long

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    
    If VarType(ap) = 8 Then
       howLong = Len(ap)
#If Win16 Then
       If howLong > 32767 Then
          Err.Clear
          myerrstr = "Third param string length must be <= 32767"
          Err.Description = myerrstr
          MsgBox myerrstr
          Err.Raise (I_ERR_PARAM) 'Raise the error
       End If
#End If
    Else
       howLong = 0
    End If
      
    ' Call the function in the SICL DLL and check for errors
    If IsMissing(ap) Then
       id = vb_ivprintf(ByVal id1, ByVal fmt, ByVal sEmpty, ByVal howLong)
    ElseIf IsEmpty(ap) Then
       id = vb_ivprintf(ByVal id1, ByVal fmt, ByVal sEmpty, ByVal howLong)
    Else
       id = vb_ivprintf(ByVal id1, ByVal fmt, ByVal ap, ByVal howLong)
    End If
    
    ivprintf = id

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function


Function ivscanf(ByVal id1 As Integer, ByVal fmt As String, ByRef myVal As Variant) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim howLong As Long
    Dim myerrstr As String * 60
    Dim tmp As Integer
    Dim returnStr As String
        
    ' Put anything in the local string to make non-null
    returnStr = "aa"
                   
    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    tmp = VarType(myVal)

    If VarType(myVal) = 8 Then
       howLong = Len(myVal)
#If Win16 Then
       If howLong > 32767 Then
          Err.Clear
          myerrstr = "Third param string length must be <= 32767"
          Err.Description = myerrstr
          MsgBox myerrstr
          Err.Raise (I_ERR_PARAM) 'Raise the error
       End If
#End If
    Else
       howLong = 0
    End If

    ' Call the function in the SICL DLL and check for errors
    If tmp = 8 Then
       id = vb_ivscanf(id1, fmt, returnStr, howLong)

       'Place scanf value into myVal
       myVal = returnStr
    Else
       id = vb_ivscanf(id1, fmt, myVal, howLong)
    End If

        
    ivscanf = id
        
    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        myerrstr = ""
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        MsgBox myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function


Function iflush(ByVal id1 As Integer, ByVal mask As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iflush(id1, mask)
    iflush = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function isetbuf(ByVal id1 As Integer, ByVal mask As Integer, ByVal size As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_isetbuf(id1, mask, size)
    isetbuf = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iclear(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iclear(id1)
    iclear = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilocal(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilocal(id1)
    ilocal = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iremote(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iremote(id1)
    iremote = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ireadstb(ByVal id1 As Integer, ByRef stb As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ireadstb(id1, stb)
    ireadstb = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function itrigger(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_itrigger(id1)
    itrigger = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ixtrig(ByVal id1 As Integer, ByVal which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ixtrig(id1, which)
    ixtrig = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function ihint(ByVal id1 As Integer, ByVal hint As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ihint(id1, hint)
    ihint = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function isetstb(ByVal id1 As Integer, ByVal stb As Byte) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_isetstb(id1, stb)
    isetstb = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ilock(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilock(id1)
    ilock = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iunlock(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iunlock(id1)
    iunlock = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function isetlockwait(ByVal id1 As Integer, ByVal flag As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_isetlockwait(id1, flag)
    isetlockwait = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetlockwait(ByVal id1 As Integer, flag As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetlockwait(id1, flag)
    igetlockwait = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function itimeout(ByVal id1 As Integer, ByVal tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_itimeout(id1, tval)
    itimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igettimeout(ByVal id1 As Integer, tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igettimeout(id1, tval)
    igettimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetaddr(ByVal id1 As Integer, ByRef addr As String) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetaddr(id1, addr)
    igetaddr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetintftype(ByVal id1 As Integer, pdata As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetintftype(id1, pdata)
    igetintftype = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetsesstype(ByVal id1 As Integer, pdata As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetsesstype(id1, pdata)
    igetsesstype = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetdevaddr(ByVal id1 As Integer, prim As Integer, sec As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetdevaddr(id1, prim, sec)
    igetdevaddr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function igetlu(ByVal id1 As Integer, lu As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetlu(id1, lu)
    igetlu = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ibeswap(addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iswap(addr, length, datasize)
    ibeswap = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function ileswap(addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
   ' We are already LE, so no swapping necesary...
   ileswap = I_ERR_NOERROR
End Function


Function iswap(addr As Variant, ByVal length As Long, ByVal datasize As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iswap(addr, length, datasize)
    iswap = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igetlulist(list() As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetlulist(list)
    igetlulist = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igetluinfo(ByVal lu As Integer, result As lu_info) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    Dim tempLu As lu_info

    tempLu.hwarg(0) = "abc0"
    tempLu.hwarg(1) = "efg1"
    tempLu.hwarg(2) = "ijk2"

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetluinfo(lu, tempLu)
    igetluinfo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    Else
       'No error, so copy data to struct
       result = tempLu
    End If

End Function


Function igetgatewaytype(ByVal id1 As Integer, pdata As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igetgatewaytype(id1, pdata)
    igetgatewaytype = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function iserialmclstat(ByVal id1 As Integer, ByVal sline As Integer, state As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialmclstat(id1, sline, state)
    iserialmclstat = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialmclctrl(ByVal id1 As Integer, ByVal sline As Integer, ByVal state As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialmclctrl(id1, sline, state)
    iserialmclctrl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialctrl(ByVal id1 As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialctrl(id1, request, setting)
    iserialctrl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialstat(ByVal id1 As Integer, ByVal request As Integer, result As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialstat(id1, request, result)
    iserialstat = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function



Function iserialbreak(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iserialbreak(id)
    iserialbreak = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxibusstatus(ByVal id1 As Integer, ByVal request As Integer, result As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxibusstatus(id1, request, result)
    ivxibusstatus = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxiwaitnormop(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxiwaitnormop(id)
    ivxiwaitnormop = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxitrigon(ByVal id1 As Integer, ByVal which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxitrigon(id1, which)
    ivxitrigon = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxitrigoff(ByVal id1 As Integer, ByVal which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxitrigoff(id1, which)
    ivxitrigoff = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxitrigroute(ByVal id1 As Integer, ByVal in_which As Long, ByVal out_which As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxitrigroute(id1, in_which, out_which)
    ivxitrigroute = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxigettrigroute(ByVal id1 As Integer, ByVal which As Long, route As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxigettrigroute(id1, which, route)
    ivxigettrigroute = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxiws(ByVal id1 As Integer, ByVal wscmd As Integer, wsresp As Integer, rpe As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxiws(id1, wscmd, wsresp, rpe)
    ivxiws = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxiservants(ByVal id1 As Integer, ByVal maxnum As Integer, list() As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxiservants(id1, maxnum, list)
    ivxiservants = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ivxirminfo(ByVal id1 As Integer, ByVal laddr As Integer, ByRef info As vxiinfo) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ivxirminfo(id1, laddr, info)
    ivxirminfo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpibbusstatus(ByVal id1 As Integer, ByVal request As Integer, result As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibbusstatus(id1, request, result)
    igpibbusstatus = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibppoll(ByVal id1 As Integer, result As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibppoll(id1, result)
    igpibppoll = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibppollconfig(ByVal id1 As Integer, ByVal cval As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibppollconfig(id1, cval)
    igpibppollconfig = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibppollresp(ByVal id1 As Integer, ByVal sval As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibppollresp(id1, sval)
    igpibppollresp = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibpassctl(ByVal id1 As Integer, ByVal busaddr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibpassctl(id1, busaddr)
    igpibpassctl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibrenctl(ByVal id1 As Integer, ByVal ren As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibrenctl(id1, ren)
    igpibrenctl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibatnctl(ByVal id1 As Integer, ByVal atnval As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibatnctl(id1, atnval)
    igpibatnctl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibsendcmd(ByVal id1 As Integer, ByVal buf As String, ByVal length As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibsendcmd(id1, buf, length)
    igpibsendcmd = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibllo(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibllo(id)
    igpibllo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibbusaddr(ByVal id1 As Integer, ByVal busaddr As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibbusaddr(id1, busaddr)
    igpibbusaddr = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibgett1delay(ByVal id1 As Integer, delay As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibgett1delay(id1, delay)
    igpibgett1delay = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function igpibsett1delay(ByVal id1 As Integer, ByVal delay As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibsett1delay(id1, delay)
    igpibsett1delay = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpibpulseifc(ByVal id1 As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpibpulseifc(id)
    igpibpulseifc = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpioctrl(ByVal id1 As Integer, ByVal request As Integer, ByVal setting As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpioctrl(id1, request, setting)
    igpioctrl = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpiostat(ByVal id1 As Integer, ByVal request As Integer, ByRef result As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpiostat(id1, request, result)
    igpiostat = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpiosetwidth(ByVal id1 As Integer, ByVal width As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpiosetwidth(id1, width)
    igpiosetwidth = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function igpiogetwidth(ByVal id1 As Integer, ByRef width As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_igpiogetwidth(id1, width)
    igpiogetwidth = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilantimeout(ByVal id1 As Integer, ByVal tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilantimeout(id1, tval)
    ilantimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilangettimeout(ByVal id1 As Integer, tval As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilangettimeout(id1, tval)
    ilangettimeout = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function imap(ByVal id1 As Integer, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer, ByVal suggested As Long) As Long
    Dim id As Long
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL and check for errors
    id = vb_imap(id1, mapspace, pagestart, pagecnt, suggested)
    imap = id
    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Function imapx(ByVal id1 As Integer, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Long
    Dim retVal As Long
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_imapx(id1, mapspace, pagestart, pagecnt)
    imapx = retVal
    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Function iunmap(ByVal id1 As Integer, ByVal addr As Long, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iunmap(id1, addr, mapspace, pagestart, pagecnt)
    iunmap = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iunmapx(ByVal id1 As Integer, ByVal addr As Long, ByVal mapspace As Integer, ByVal pagestart As Integer, ByVal pagecnt As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iunmapx(id1, addr, mapspace, pagestart, pagecnt)
    iunmapx = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function imapinfo(ByVal id1 As Integer, ByVal mapspace As Integer, numwindows As Integer, winsize As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_imapinfo(id1, mapspace, numwindows, winsize)
    imapinfo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function ipokex8(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal offset As Long, ByVal Value As Byte) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipokex8(siclId, mapHandle, offset, Value)
    ipokex8 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipeekx8(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal offset As Long, Value As Byte) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipeekx8(siclId, mapHandle, offset, Value)
    ipeekx8 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipokex16(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal offset As Long, ByVal Value As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipokex16(siclId, mapHandle, offset, Value)
    ipokex16 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipeekx16(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal offset As Long, Value As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipeekx16(siclId, mapHandle, offset, Value)
    ipeekx16 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipokex32(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal offset As Long, ByVal Value As Long) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipokex32(siclId, mapHandle, offset, Value)
    ipokex32 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ipeekx32(ByVal siclId As Integer, ByVal mapHandle As Long, ByVal offset As Long, Value As Long) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_ipeekx32(siclId, mapHandle, offset, Value)
    ipeekx32 = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function iblockmovex(ByVal siclId As Integer, ByVal srcHandle As Long, ByRef srcOffset As Variant, ByVal srcWidth As Integer, ByVal srcIncrement As Integer, ByVal destHandle As Long, ByRef destOffset As Variant, ByVal destWidth As Integer, ByVal destIncrement As Integer, ByVal count As Long, ByVal swap As Integer) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iblockmovex(siclId, srcHandle, srcOffset, srcWidth, srcIncrement, destHandle, destOffset, destWidth, destIncrement, count, swap)
    iblockmovex = retVal
    If retVal <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If
End Function

Function ibblockcopy(ByVal id1 As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibblockcopy(id1, src, dest, cnt)
    ibblockcopy = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwblockcopy(ByVal id1 As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwblockcopy(id1, src, dest, cnt, swap)
    iwblockcopy = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilblockcopy(ByVal id1 As Integer, ByVal src As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilblockcopy(id1, src, dest, cnt, swap)
    ilblockcopy = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ibpushfifo(ByVal id1 As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibpushfifo(id1, src, fifo, cnt)
    ibpushfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwpushfifo(ByVal id1 As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwpushfifo(id1, src, fifo, cnt, swap)
    iwpushfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilpushfifo(ByVal id1 As Integer, ByVal src As Long, ByVal fifo As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilpushfifo(id1, src, fifo, cnt, swap)
    ilpushfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ibpopfifo(ByVal id1 As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibpopfifo(id1, fifo, dest, cnt)
    ibpopfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function iwpopfifo(ByVal id1 As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwpopfifo(id1, fifo, dest, cnt, swap)
    iwpopfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

Function ilpopfifo(ByVal id1 As Integer, ByVal fifo As Long, ByVal dest As Long, ByVal cnt As Long, ByVal swap As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilpopfifo(id1, fifo, dest, cnt, swap)
    ilpopfifo = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function siclcleanup() As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb__siclcleanup()
    siclcleanup = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Function setsiclyield(ByVal yield_option As Integer) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb__setsiclyield(yield_option)
    setsiclyield = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function


Sub ibpoke(ByVal addr As Long, ByVal Value As Byte)
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL
    Call vb_ibpoke(addr, Value)

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Sub


Sub iwpoke(ByVal addr As Long, ByVal Value As Integer)
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL
    Call vb_iwpoke(addr, Value)

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Sub

Sub ilpoke(ByVal addr As Long, ByVal Value As Long)
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Force no error to be condition
    Call vb_icauseerr(id1, 0, 0)

    ' Call the function in the SICL DLL
    Call vb_ilpoke(addr, Value)

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Sub

Function ibpeek(ByVal addr As Long) As Byte
    Dim id As Byte

    ' Call the function in the SICL DLL and check for errors
    id = vb_ibpeek(addr)
    ibpeek = id

End Function

Function iwpeek(ByVal addr As Long) As Integer
    Dim id As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_iwpeek(addr)
    iwpeek = id

End Function

Function ilpeek(ByVal addr As Long) As Long
    Dim id As Long

    ' Call the function in the SICL DLL and check for errors
    id = vb_ilpeek(addr)
    ilpeek = id

End Function


' This function truncates a string so that all characters
' following a carriage return or linefeed character are
' removed.  The truncated string is then returned.
Function strip_crlf(szString As String) As String
   Dim crlfpos As Integer

   crlfpos = InStr(szString, Chr$(13))
   If crlfpos Then
     szString = Left(szString, crlfpos - 1)
   End If
   crlfpos = InStr(szString, Chr$(10))
   If crlfpos Then
     szString = Left(szString, crlfpos - 1)
   End If

   strip_crlf = szString
End Function

Function icmd(ByVal id1 As Integer, ByVal cmd As Long, ByVal datalen As Integer, ByVal datawidth As Integer, ByRef pdata As Long) As Integer
    Dim id As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer

    ' Call the function in the SICL DLL and check for errors
    id = vb_icmd(id1, cmd, datalen, datawidth, pdata)
    icmd = id
    If id <> 0 Then
        thisErrno = vb_igeterrno()
        If thisErrno <> 0 Then
            Err.Clear    ' set default values in the error object
            ' set the error string and raise the error
            tmp = vb_igeterrstr(thisErrno, myerrstr)
            Err.Description = myerrstr
            Err.Raise (thisErrno) 'Raise the error
        End If
    End If

End Function

#If Win32 Then
Public Function isprintf(ByRef s As String, ByVal fmt As String, ParamArray vararg() As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    If (UBound(vararg) <> -1) Then
       ReDim va(LBound(vararg) To UBound(vararg)) As Variant
    Else
       ReDim va(0 To 0) As Variant
       va(0) = 0
    End If
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
    
    For i = LBound(vararg) To UBound(vararg)
       va(i) = vararg(i)
    Next i

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iprint(1, 0, s, ByVal fmt, va)
    
    isprintf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Public Function iprintf(ByRef id As Integer, ByVal fmt As String, ParamArray vararg() As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    If (UBound(vararg) <> -1) Then
       ReDim va(LBound(vararg) To UBound(vararg)) As Variant
    Else
       ReDim va(0 To 0) As Variant
       va(0) = 0
    End If
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
    
    For i = LBound(vararg) To UBound(vararg)
       va(i) = vararg(i)
    Next i

    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iprint(2, id, 0, ByVal fmt, va)
    
    iprintf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Public Function isscanf(ByVal s As String, ByVal fmt As String, Optional va1 As Variant, Optional va2 As Variant, Optional va3 As Variant, Optional va4 As Variant, Optional va5 As Variant, Optional va6 As Variant, Optional va7 As Variant, Optional va8 As Variant, Optional va9 As Variant, Optional va10 As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
    
    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iscan(3, 0, ByVal s, ByVal fmt, va1, va2, va3, va4, va5, va6, va7, va8, va9, va10)
    
    isscanf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

Public Function iscanf(ByRef id As Integer, ByVal fmt As String, Optional va1 As Variant, Optional va2 As Variant, Optional va3 As Variant, Optional va4 As Variant, Optional va5 As Variant, Optional va6 As Variant, Optional va7 As Variant, Optional va8 As Variant, Optional va9 As Variant, Optional va10 As Variant) As Integer
    Dim retVal As Integer
    Dim thisErrno As Integer
    Dim myerrstr As String * 60
    Dim tmp As Integer
    
    ' Force no error to be condition
    Call vb_icauseerr(0, 0, 0)
   
    ' Call the function in the SICL DLL and check for errors
    retVal = vb_iscan(4, id, 0, ByVal fmt, va1, va2, va3, va4, va5, va6, va7, va8, va9, va10)
    
    iscanf = retVal

    thisErrno = vb_igeterrno()
    If thisErrno <> 0 Then
        Err.Clear    ' set default values in the error object
        ' set the error string and raise the error
        tmp = vb_igeterrstr(thisErrno, myerrstr)
        Err.Description = myerrstr
        Err.Raise (thisErrno) 'Raise the error
    End If

End Function

#End If
    
