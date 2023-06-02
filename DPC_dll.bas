Attribute VB_Name = "Dpc_dll"
Option Explicit

Type DigitalProductId
   CONST_Size As Long
   CONST_VER As Long
   ProductID As String * &H18
   BINK_ID As Long               '//+20 =signaturID+SomeBitFlag
   StockID As String * &H10
   CDKEY As String * &H14 'byte
   CreationDate As Long  'time_t
   Year_and_more As Long
   PRC_Mode As Long 'RPC_TYPE
   Unused2 As Long
   String1 As String * &H10
   String2 As String * &HC
   Expire_Year As Long
   BLINK_Date As Long   'time_t
   Value1_L3 As Long
   Value2 As Long
   String3 As String * &H1C
   Checksum As Long
End Type
Public DigitalProductId As DigitalProductId

Declare Function ValidateDigitalPid& Lib "dpcdll" Alias "#123" (DigitalPID As Any)
Declare Function CheckCDKey& Lib "dpcdll" Alias "#125" _
   (ByVal Pid30Text$, ByVal Pid30Rpc$, ByVal StockID$, _
   ByVal SomeString$, ByVal RESERVED&, ByVal RESERVED&, _
   ByVal RESERVED&, ByVal BINK_ID_Flag&, ByVal ProductID$, _
   DigitalProductId As DigitalProductId, ByVal RESERVED&, ByVal RESERVED&)

Public Const CP_ACP As Long = 0
Public Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long

Function uni$(ByVal acci$)
   Dim RetVal&, tmpstr$
   tmpstr = Space(Len(acci) * 2)
   RetVal = MultiByteToWideChar(CP_ACP, 0, acci, -1, tmpstr, Len(acci) + 1)
   uni = tmpstr
End Function

