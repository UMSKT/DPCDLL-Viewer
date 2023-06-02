VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (src As Any, src As Any, ByVal Length&)
Private Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal src As String, ByVal Length&)
Private Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal src As String, src As Long, ByVal Length&)
Private Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal src As Integer, ByVal Length&)

Public bSearchBackward As Boolean
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarPosition As Long 'lokale Kopie
Private mvardata As String 'lokale Kopie
'Public Position As Long 'lokale Kopie
Public mStorePos As Long
Public DisableAutoMove As Boolean
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarLength As Long 'lokale Kopie
Public Property Let Length(ByVal vData As Long)
    Stop 'not implemented
    mvarLength = vData
End Property


Public Property Get Length() As Long
    Length = Len(mvardata)
End Property




Public Sub RestorePos()
   Position = mStorePos
End Sub

Public Sub StorePos()
   mStorePos = Position
End Sub


Public Property Let ToMove(ByVal vData As Long)
   If DisableAutoMove Then Exit Property
   Move (vData)
End Property


Public Property Let Position(vData As Long)
    mvarPosition = limit(vData, Len(mvardata), 0) + 1
End Property


Public Property Get Position() As Long
    Position = mvarPosition - 1
End Property


Private Function StrtoLng(ByVal value$) As Long
   MemCopyStrToLng StrtoLng, value, 4
End Function
'
Private Function LngtoStr(ByVal value&) As String
   Dim tmp$
   tmp = Space(4)
   MemCopyLngToStr tmp, value, 4
   LngtoStr = tmp
End Function

'/////////////////////////////////////////////////////////
'// set_EOS - Returns True if Position is at the End Of String
Public Property Let EOS(ByVal vData As Boolean)
    Position = Len(mvardata) And vData
End Property
'// get_EOS - Forward to End Of String
Public Property Get EOS() As Boolean
   EOS = Position >= Len(mvardata)
End Property


Public Sub Move(Chars&)
   Position = Position + Chars
'   Debug.Print "Move: ", Chars
End Sub

Public Function FindByte(ByVal LongValue_To_Find$, Optional Range) As Long
   FindByte = Findstring(Left(LngtoStr(LongValue_To_Find), 1))
End Function

Public Function FindInt(ByVal LongValue_To_Find$, Optional Range) As Long
   FindInt = Findstring(Left(LngtoStr(LongValue_To_Find), 2))
End Function

Public Function FindLong(ByVal LongValue_To_Find$, Optional Range) As Long
   FindLong = Findstring(LngtoStr(LongValue_To_Find))
End Function

Public Function Findstring(String_To_Find$, Optional Range) ', Optional Alternativ_String_To_Find) As Long
   If IsMissing(Range) Then Range = Len(mvardata)
   
'   Findstring = InStr(1, Mid(mvarData, mvarPosition, Range), String_To_Find)
   If bSearchBackward Then
      Findstring = InStrRev(mvardata, String_To_Find, mvarPosition)
      
     'Test if out of range
      If (mvarPosition - Findstring) > Range Then
         Findstring = 0
      End If

   
   Else
      Findstring = InStr(mvarPosition, mvardata, String_To_Find)
      
     'Test if out of range
      If (Findstring - mvarPosition) > Range Then
         Findstring = 0
      End If

   End If
      
  
   
'   If IsMissing(Alternativ_String_To_Find = False) And (Findstring = 0) Then
'   Findstring = InStr(1, Mid(mvarData, Position, Range), String_To_Find)
   
 
 ' If string was found
   If Findstring Then
     
     'Return start of String
'      Findstring = (Findstring - 1) + Position
      Findstring = (Findstring - 1)
 
     
     'seek at the end of found String
      ToMove = (Findstring - Position) + Len(String_To_Find)

     
   End If
End Function

Public Property Let FixedString(Optional ByVal Length& = -1, vData As String)
    If Length <= 0 Then Length = Len(vData)
    If Length <= 0 Then Exit Property
    
   'Enlarge Buffer if necessary
    Dim enlarge&
    enlarge = (Length + Position) - Len(mvardata)
    If enlarge >= 1 Then mvardata = mvardata & Space(enlarge)
    
    Mid(mvardata, Position + 1, Length) = vData
    ToMove = Length
End Property


Public Property Get FixedString(Optional ByVal Length& = -1) As String
    If Length <= 0 Then Length = Len(mvardata)
    FixedString = Mid(mvardata, mvarPosition, Length)
   
    ToMove = Length
End Property



Public Property Let int32(vData As Long)
    FixedString = LngtoStr(vData)
End Property


Public Property Get int32() As Long
    int32 = StrtoLng(FixedString(4))
End Property



Public Property Let int16(vData As Long)
    FixedString(2) = LngtoStr(vData)
End Property


Public Property Get int16() As Long
    int16 = StrtoLng(FixedString(2))
End Property



Public Property Let int8(vData As Long)
    FixedString(1) = LngtoStr(vData)
End Property


Public Property Get int8() As Long
    int8 = StrtoLng(FixedString(1))
End Property




Public Property Let data(vData As String)
    mvardata = vData
    Position = 0
End Property


Public Property Get data() As String
Attribute data.VB_UserMemId = 0
    data = mvardata
End Property

