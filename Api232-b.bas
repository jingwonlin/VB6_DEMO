Attribute VB_Name = "ModApi232"

'       api232-b.bas      ver 1.00
'       API232 define file for Win32 on Visual Basic
'       History:
'       Date        Author      Comment
'       01/10/1997  Victor      wrote it.
'       03/28/1997  Victor      modified. add sio_linput funciton
'       04/29/1997  Victor      modified. add sio_putb_x function
'       06/30/1997  Victor      modified. add sio_TxLowWater function
'       11/17/1997  Victor      modified. add sio_SetWriteTimeouts,
'                                             sio_AbortWrite
'       11/24/1997  Victor      add sio_SetReadTimeouts, sio_GetReadTimeouts
'                               sio_GetWriteTimeouts, sio_AbortRead

'       Baud Rate Setting
Global Const B50 = &H0
Global Const B75 = &H1
Global Const B110 = &H2
Global Const B134 = &H3
Global Const B150 = &H4
Global Const B300 = &H5
Global Const B600 = &H6
Global Const B1200 = &H7
Global Const B1800 = &H8
Global Const B2400 = &H9
Global Const B4800 = &HA
Global Const B7200 = &HB
Global Const B9600 = &HC
Global Const B19200 = &HD
Global Const B38400 = &HE
Global Const B57600 = &HF
Global Const B115200 = &H10
Global Const B230400 = &H11
Global Const B460800 = &H12
Global Const B921600 = &H13

'       MODE setting
Global Const BIT_5 = &H0                             ' Word length define
Global Const BIT_6 = &H1
Global Const BIT_7 = &H2
Global Const BIT_8 = &H3

Global Const STOP_1 = &H0                            ' Stop bits define
Global Const STOP_2 = &H4

Global Const P_EVEN = &H18                           ' Parity define
Global Const P_ODD = &H8
Global Const P_SPC = &H38
Global Const P_MRK = &H28
Global Const P_NONE = &H0

'       MODEM CONTROL setting
Global Const C_DTR = &H1
Global Const C_RTS = &H2

'       MODEM LINE STATUS
Global Const S_CTS = &H1
Global Const S_DSR = &H2
Global Const S_RI = &H4
Global Const S_CD = &H8


'  error code
Global Const SIO_OK = 0
Global Const SIO_BADPORT = -1        ' no such port or port not opened
Global Const SIO_OUTCONTROL = -2     ' can't control MOXA board
Global Const SIO_NODATA = -4         ' no data to read or no buffer to write
Global Const SIO_BADPARM = -7        ' bad parameter
Global Const SIO_WIN32FAIL = -8      ' call win32 function fail, please call
                                     ' GetLastError to get the error code
Global Const CODE = &HFF
Global Const DLE = &H10
Global Const STX = &H2
Global Const ETX = &H3
Global Const ACK = &H6
Global Const NAK = &H15
Global Const ENQ = &H5
Global Const EOT = &H4



Declare Function sio_ioctl Lib "api232.dll" (ByVal port As Long, ByVal Baud As Long, ByVal mode As Long) As Long
Declare Function sio_getch Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_read Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal length As Long) As Long
Declare Function sio_putch Lib "api232.dll" (ByVal port As Long, ByVal term As Long) As Long
Declare Function sio_write Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal length As Long) As Long
Declare Function sio_flush Lib "api232.dll" (ByVal port As Long, ByVal func As Long) As Long
Declare Function sio_iqueue Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_oqueue Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_lstatus Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_lctrl Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_break Lib "api232.dll" (ByVal port As Long, ByVal time As Long) As Long
Declare Function sio_flowctrl Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_Tx_hold Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_close Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_open Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_getbaud Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_getmode Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_getflow Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_DTR Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_RTS Lib "api232.dll" (ByVal port As Long, ByVal mode As Long) As Long
Declare Function sio_baud Lib "api232.dll" (ByVal port As Long, ByVal speed As Long) As Long
Declare Function sio_data_status Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_putb Lib "api232.dll" Alias "sio_write" (ByVal port As Long, ByVal buf As String, ByVal length As Long) As Long
Declare Function sio_linput Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal length As Long, ByVal term As Long) As Long
Declare Function sio_putb_x Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal length As Long, ByVal tick As Long) As Long
Declare Function sio_view Lib "api232.dll" (ByVal port As Long, ByVal buf As String, ByVal length As Long) As Long
Declare Function sio_TxLowWater Lib "api232.dll" (ByVal port As Long, ByVal size As Long) As Long
Declare Function sio_SetWriteTimeouts Lib "api232.dll" (ByVal port As Long, ByVal timeouts As Long) As Long
Declare Function sio_AbortWrite Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_SetReadTimeouts Lib "api232.dll" (ByVal port As Long, ByVal TotalTimeouts As Long, ByVal IntervalTimeouts As Long) As Long
Declare Function sio_AbortRead Lib "api232.dll" (ByVal port As Long) As Long
Declare Function sio_GetReadTimeouts Lib "api232.dll" (ByVal port As Long, ByRef TotalTimeouts As Long, ByRef InterfalTimeouts As Long) As Long
Declare Function sio_GetWriteTimeouts Lib "api232.dll" (ByVal port As Long, ByRef TotalTimeouts As Long) As Long

'Moxa card port open
Public Function moxaOpen(iCom As Long, Ibaud As Long, Sparity As String, Idatabit As Long, Istopbit As Long) As Long

' Open the serial port.
' Icom means to it is Moxa port number
' Set the port baud rate is Ibaud and Iparity,Idatabit,Istopbit by sio_ioctl.
'return code = -1  setting value error
'              -2  setting error
 
 Dim BaudLong, BitLong, StopLong, ParityLong As Long
    moxaOpen = sio_open(iCom)
    If moxaOpen < 0 Then
        Exit Function
    End If
    If Ibaud = 2400 Then
        BaudLong = B2400
    ElseIf Ibaud = 4800 Then
        BaudLong = B4800
    ElseIf Ibaud = 9600 Then
        BaudLong = B9600
    ElseIf Ibaud = 19200 Then
        BaudLong = B19200
    ElseIf Ibaud = 38400 Then
        BaudLong = B38400
    Else
        moxaOpen = -2
        Exit Function
    End If
    ':set databit
    If Idatabit = 8 Then
        BitLong = BIT_8
    ElseIf Idatabit = 7 Then
        BitLong = BIT_7
    Else
        moxaOpen = -2
        Exit Function
    End If
    
    ':set StopBit
    If Istopbit = 1 Then
        StopLong = &H0
    ElseIf Istopbit = 2 Then
        StopLong = &H4
    Else
        moxaOpen = -2
        Exit Function
    End If
    ':Parity
    If Sparity = "EVEN" Then
        ParityLong = P_EVEN
    ElseIf Sparity = "ODD" Then
        ParityLong = P_ODD
    ElseIf Sparity = "NONE" Then
        ParityLong = P_NONE
    Else
        moxaOpen = -2
        Exit Function
    End If
    
    moxaOpen = sio_ioctl(iCom, BaudLong, BitLong + StopLong + ParityLong)
End Function


