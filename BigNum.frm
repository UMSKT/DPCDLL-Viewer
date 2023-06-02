VERSION 5.00
Begin VB.Form BigNum 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Thanks to Andrija Radovic for this program and its sourcecode !"
   ClientHeight    =   5385
   ClientLeft      =   555
   ClientTop       =   1155
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BigNum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9375
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ComboBox combo_registry 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "BigNum.frx":0442
      Left            =   2160
      List            =   "BigNum.frx":0452
      TabIndex        =   17
      ToolTipText     =   "Note: You may need to edit the path to 'DigitalProductID' if you get and error"
      Top             =   3840
      Width           =   7095
   End
   Begin VB.TextBox txt_BINK 
      Height          =   390
      Left            =   4440
      TabIndex        =   16
      Text            =   "_2"
      ToolTipText     =   "dd"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txt_Site 
      Height          =   390
      Left            =   2880
      TabIndex        =   15
      Text            =   "_47"
      ToolTipText     =   "bbb"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txt_String 
      Height          =   390
      Left            =   4440
      TabIndex        =   14
      Text            =   "    "
      ToolTipText     =   "SomeAdditionData"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txt_sku 
      Height          =   390
      Left            =   2040
      TabIndex        =   13
      Text            =   "_22-00001"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txt_RPC 
      Height          =   390
      Left            =   2040
      TabIndex        =   12
      Text            =   "_5375"
      ToolTipText     =   "aaaaa"
      Top             =   2520
      Width           =   810
   End
   Begin VB.TextBox txt_Pid 
      Height          =   390
      Left            =   2040
      TabIndex        =   11
      Text            =   "_5375-647-1979925-22376"
      Top             =   2040
      Width           =   5175
   End
   Begin VB.CommandButton cmd_SetCDKey 
      Caption         =   "Set CDKey = "
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmd_get_cdkey 
      Caption         =   "Get key"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Reads the CDKey from the Registry('DigitalProductID')"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extract"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Extracts the -bbb- and -ccccccc- part of the ProductID"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Txt_CDKey 
      Height          =   390
      Left            =   2040
      TabIndex        =   0
      Text            =   "JY7PB-RCQY7-H6JWB-Q7HKC-W3PGG"
      ToolTipText     =   "! = FACT, ~ = NOT, @ = XOR, & = AND, % = OR"
      Top             =   1560
      Width           =   7095
   End
   Begin VB.TextBox Txt_Bin 
      Height          =   390
      Left            =   2040
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Operators: X+Y, X-Y, X%Y, X&Y, X@Y, X/Y, X*Y, X^Y, ~X, X!"
      Top             =   1080
      Width           =   7095
   End
   Begin VB.TextBox Txt_Dec 
      Height          =   390
      Left            =   2040
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Example: 5!*2/3"
      Top             =   600
      Width           =   7095
   End
   Begin VB.TextBox Txt_Hex 
      Height          =   390
      Left            =   2040
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Example: AF*5+2"
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label PID 
      Alignment       =   1  'Rechts
      Caption         =   "Bin PID"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   22
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label PID 
      Alignment       =   1  'Rechts
      Caption         =   "Text PID"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   21
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label sku 
      Alignment       =   1  'Rechts
      Caption         =   "Text SKU"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lbl_Status 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   4440
      Width           =   9375
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Rechts
      Caption         =   "Registrypath"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Rechts
      Caption         =   "CDKey = "
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Rechts
      Caption         =   "BIN = "
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Rechts
      Caption         =   "DEC ="
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Rechts
      Caption         =   "HEX = "
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu ExitAll 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "BigNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Autor: Andrija Radovic
Const Dec As String = "0123456789"
Const Hex As String = "0123456789ABCDEF"
Const Oct As String = "BCDFGHJKMPQRTVWXY2346789"
'Const Oct As String = "3456789BCDFGHJKLMNPQRSTVWXY"
'Const Oct As String = "01234567"
Const Bin As String = "01"
Dim reg As New Registry
Dim RetVal&

Private Sub SetValues()
 ' Issue new DigitalProductId
   DigitalProductId.CONST_Size = Len(DigitalProductId) '&HA4
   RetVal = CheckCDKey(uni(Txt_CDKey), uni(txt_RPC), uni(txt_sku), uni(txt_String), 0, 0, 0, 0, uni(txt_Pid), DigitalProductId, 0, 0)
   If RetVal <> 0 Then RetVal = CheckCDKey(uni(Txt_CDKey), uni(txt_RPC), uni(txt_sku), uni("    "), 0, 0, 0, 1, uni(txt_Pid), DigitalProductId, 0, 0)
   If RetVal = 0 Then
    ' RetVal = CheckCDKey(uni("X3WJB-3B2BH-3MPM6-8F6GR-X9HBJ"), uni("55375"), uni("A22-00001"), uni("    "), 0, 0, 0, 0, uni("55375-013-4130274-22293"), DigitalProductId, 0, 0)
      RetVal = ValidateDigitalPid(DigitalProductId)
      If RetVal = 0 Then
         Dim tmp$
         tmp = Space(Len(DigitalProductId))
         MemCopyRefToRef ByVal tmp, DigitalProductId, Len(DigitalProductId)
         reg.Regdata = tmp
         
         GetValues
      Else
         MsgBox "ValidateDigitalPid Errorcode: " & RetVal, vbCritical, "DigitalProductId was not set"
      End If
   Else
      MsgBox "CheckCDKey Errorcode: " & RetVal, vbCritical, "DigitalProductId was not set"
   End If
   
   '"GJQHW-B4KYW-KGWBR-B3FR6-TB34M"
   
End Sub


Private Sub GetValues()

   
   MemCopyRefToRef DigitalProductId, ByVal CStr(reg.Regdata), Len(DigitalProductId)
   With DigitalProductId
      txt_Pid = .ProductID
      txt_sku = .StockID
      txt_String = Mid(.String1, 5, 4)
      txt_RPC = Split(.ProductID, "-")(0)
      txt_Site = Split(.ProductID, "-")(1)
      txt_BINK = .BINK_ID
   End With
   
   
'   DigitalProductId.CONST_Size = Len(DigitalProductId) '&HA4
   
   'RetVal = CheckCDKey(uni("X3WJB-3B2BH-3MPM6-8F6GR-X9HBJ"), uni("55375"), uni("A22-00001"), uni("    "), 0, 0, 0, 0, uni("55375-013-4130274-22293"), DigitalProductId, 0, 0)
   
'   RetVal = CheckCDKey(uni("X3WJB-3B2BH-3MPM6-8F6GR-X9HBJ"), uni("     "), uni("         "), uni("    "), 0, 0, 0, 0, 0, tmpstr, 0, 0)

'   RetVal = CheckCDKey(uni("GJQHW-B4KYW-KGWBR-B3FR6-TB34M"), uni("55375"), uni("A22-00001"), uni("    "), 0, 0, 0, 0, uni("55375-013-4130274-22293"), tmpstr, 0, 0)
   '"GJQHW-B4KYW-KGWBR-B3FR6-TB34M"

End Sub

Private Sub cmd_get_cdkey_Click()

   
   Txt_Hex = ""
   Dim i
'   For i = &H43 To &H35 Step -1
'      Txt_Hex = Txt_Hex & H8(Asc(Mid(DPID, i, 1)))
'   Next
   For i = 1 To Len(DigitalProductId.CDKEY)
      Txt_Hex = H8(Asc(Mid(DigitalProductId.CDKEY, i, 1))) & Txt_Hex
   Next
      
   Txt_Hex_LostFocus
   
End Sub

Private Sub cmd_SetCDKey_Click()
   SetValues
End Sub

Private Sub combo_registry_Change()
   On Error Resume Next
   lbl_Status.Caption = "Set new user RegPath: " & combo_registry.Text
   reg.Create HKEY_LOCAL_MACHINE, combo_registry.Text
   If err Then
      lbl_Status.Caption = lbl_Status.Caption & " FAILED!" & vbCrLf & err.Description
   Else
      GetValues
      If err Then
         lbl_Status.Caption = lbl_Status.Caption & " FAILED!" & vbCrLf & err.Description
      End If
      cmd_get_cdkey_Click
      
   End If

End Sub

Private Sub combo_registry_Click()
   combo_registry_Change
End Sub


Private Sub Command1_Click()
   Txt_Bin = Left(Right(Txt_Bin, 31), 30)
   Txt_Bin_LostFocus
   Txt_Dec = Right(String(Len(Txt_Dec), "0") & Txt_Dec, 9)
End Sub

Private Sub ExitAll_Click()
    End
End Sub

Private Sub Form_Load()
   On Error Resume Next
   
   reg.RegValue = "DigitalProductId"
   
   GetValues
   
   combo_registry.ListIndex = 0
  
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Txt_Hex_GotFocus()
    GetFocus Txt_Hex
End Sub

Private Sub Txt_Hex_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        Txt_CDKey.SetFocus
    Case vbKeyDown
        Txt_Dec.SetFocus
    End Select
End Sub

Private Sub Txt_Hex_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Txt_Hex_LostFocus
        GetFocus Txt_Hex
'    Case 97 To 102
'        KeyAscii = KeyAscii - 32
'    Case vbKey0 To vbKey9, vbKeyA To vbKeyF, vbKeyBack, vbKeyExecute, 45, 42, 94, 40, 41, 33, 35, 47, 37, 38, 126, 64
'        With Txt_Hex
'            If .SelStart Then
'                If Not Sintax(Mid$(.Text, .SelStart, 1), Chr$(KeyAscii), .Text) Then
'                    KeyAscii = 0
'                End If
'            Else
'                If InStr("+%&@/*^!", Chr$(KeyAscii)) Then KeyAscii = 0
'            End If
'        End With
'    Case Else
'        KeyAscii = 0
'        Beep
    End Select
End Sub

Private Sub Txt_Hex_LostFocus()
    If Txt_Hex.Tag <> Txt_Hex.Text Then
        Screen.MousePointer = vbArrowHourglass
        Txt_Hex.Text = EVAL(Txt_Hex.Text, Txt_Hex.Tag, Hex)
        Txt_Dec.Text = HEX2DEC$(Txt_Hex.Text)
        Txt_Bin.Text = HEX2BIN$(Txt_Hex.Text)
        Txt_CDKey.Text = HEX2OCT$(Txt_Hex.Text)
        ClearTags
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Txt_Dec_GotFocus()
    GetFocus Txt_Dec
End Sub

Private Sub Txt_Dec_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        Txt_Hex.SelStart = Len(Txt_Hex.Text)
        Txt_Hex.SetFocus
    Case vbKeyDown
        Txt_Bin.SetFocus
    End Select
End Sub

Private Sub Txt_Dec_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Txt_Dec_LostFocus
        GetFocus Txt_Dec
'    Case vbKey0 To vbKey9, vbKeyBack, vbKeyExecute, 45, 42, 94, 40, 41, 33, 35, 47, 37, 38, 126, 64
'        With Txt_Dec
'            If .SelStart Then
'                If Not Sintax(Mid$(.Text, .SelStart, 1), Chr$(KeyAscii), .Text) Then
'                    KeyAscii = 0
'                End If
'            Else
'                If InStr("+%&@/*^!", Chr$(KeyAscii)) Then KeyAscii = 0
'            End If
'        End With
'    Case Else
'        KeyAscii = 0
'        Beep
    End Select
End Sub

Private Sub Txt_Dec_LostFocus()
    If Txt_Dec.Tag <> Txt_Dec.Text Then
        Screen.MousePointer = vbArrowHourglass
        Txt_Dec.Text = EVAL(Txt_Dec.Text, Txt_Dec.Tag, Dec)
        Txt_Hex.Text = DEC2HEX$(Txt_Dec.Text)
        Txt_Bin.Text = DEC2BIN$(Txt_Dec.Text)
        Txt_CDKey.Text = DEC2OCT$(Txt_Dec.Text)
        ClearTags
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Txt_Bin_GotFocus()
    GetFocus Txt_Bin
End Sub

Private Sub Txt_Bin_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        Txt_Dec.SetFocus
    Case vbKeyDown
        Txt_CDKey.SetFocus
    End Select
End Sub

Private Sub Txt_Bin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Txt_Bin_LostFocus
        GetFocus Txt_Bin
'    Case vbKey0, vbKey1, vbKeyBack, vbKeyExecute, 45, 42, 94, 40, 41, 33, 35, 47, 37, 38, 126, 64
'        With Txt_Bin
'            If .SelStart Then
'                If Not Sintax(Mid$(.Text, .SelStart, 1), Chr$(KeyAscii), .Text) Then
'                    KeyAscii = 0
'                End If
'            Else
'                If InStr("+%&@/*^!", Chr$(KeyAscii)) Then KeyAscii = 0
'            End If
'        End With
'    Case Else
'        KeyAscii = 0
'        Beep
    End Select
End Sub

Private Sub Txt_Bin_LostFocus()
    If Txt_Bin.Tag <> Txt_Bin.Text Then
        Screen.MousePointer = vbArrowHourglass
        Txt_Bin.Text = EVAL(Txt_Bin.Text, Txt_Bin.Tag, Bin)
        Txt_Hex.Text = BIN2HEX$(Txt_Bin.Text)
        Txt_Dec.Text = BIN2DEC$(Txt_Bin.Text)
        Txt_CDKey.Text = BIN2OCT$(Txt_Bin.Text)
        ClearTags
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Txt_CDKey_GotFocus()
   On Error Resume Next
    GetFocus Txt_CDKey
End Sub

Private Sub Txt_CDKey_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        Txt_Bin.SetFocus
    Case vbKeyDown
        Txt_Hex.SetFocus
    End Select
End Sub

Private Sub Txt_CDKey_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Txt_CDKey_LostFocus
        GetFocus Txt_CDKey
'    Case vbKey0 To vbKey7, vbKeyBack, vbKeyExecute, 45, 42, 94, 40, 41, 33, 35, 47, 37, 38, 126, 64
'        With Txt_CDKey
'            If .SelStart Then
'                If Not Sintax(Mid$(.Text, .SelStart, 1), Chr$(KeyAscii), .Text) Then
'                    KeyAscii = 0
'                End If
'            Else
'                If InStr("+%&@/*^!", Chr$(KeyAscii)) Then KeyAscii = 0
'            End If
'        End With
'    Case Else
'        KeyAscii = 0
'        Beep
    End Select
End Sub
Private Function RemoveInvalidChars$(ByRef s$, validchars$)
Dim i&, char$

'validchars = validchars
For i = 1 To Len(s)
   char = Mid(s, i, 1)
   If InStr(1, validchars, char, vbTextCompare) Then
      RemoveInvalidChars = RemoveInvalidChars & char
   End If
Next


End Function

Private Sub Txt_CDKey_LostFocus()
   On Error Resume Next
    If Txt_CDKey.Tag <> Txt_CDKey.Text Then
        Screen.MousePointer = vbArrowHourglass
        Txt_CDKey.Text = RemoveInvalidChars(Txt_CDKey.Text, Oct)
        Txt_CDKey.Text = EVAL(Txt_CDKey.Text, Txt_CDKey.Tag, Oct)
        Txt_Hex.Text = OCT2HEX$(Txt_CDKey.Text)
        Txt_Dec.Text = OCT2DEC$(Txt_CDKey.Text)
        Txt_Bin.Text = OCT2BIN$(Txt_CDKey.Text)
        ClearTags
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub ClearTags()
    Dim A As Variant
    For Each A In Array(Txt_Hex, Txt_Dec, Txt_Bin, Txt_CDKey)
        A.Tag = A.Text
    Next
End Sub

Function HEX2DEC(ByVal A As String) As String
    HEX2DEC = A2BS(A, Hex, Dec)
End Function

Function DEC2HEX(ByVal A$) As String
    DEC2HEX = A2BS(A$, Dec, Hex)
End Function

Function DEC2BIN(ByVal A$) As String
    DEC2BIN = A2BS(A$, Dec, Bin)
End Function

Function BIN2DEC(ByVal A$) As String
    BIN2DEC = A2BS(A$, Bin, Dec)
End Function

Function HEX2BIN(ByVal A$) As String
    HEX2BIN = A2BS(A$, Hex, Bin)
End Function

Function BIN2HEX(ByVal A$) As String
    BIN2HEX = A2BS(A$, Bin, Hex)
End Function

Function OCT2DEC(ByVal A$) As String
    OCT2DEC = A2BS(A$, Oct, Dec)
End Function

Function DEC2OCT(ByVal A$) As String
    DEC2OCT = A2BS(A$, Dec, Oct)
End Function

Function OCT2BIN(ByVal A$) As String
    OCT2BIN = A2BS(A$, Oct, Bin)
End Function

Function BIN2OCT(ByVal A$) As String
    BIN2OCT = A2BS(A$, Bin, Oct)
End Function

Function OCT2HEX(ByVal A$) As String
    OCT2HEX = A2BS(A$, Oct, Hex)
End Function

Function HEX2OCT(ByVal A$) As String
    HEX2OCT = A2BS(A$, Hex, Oct)
End Function

Private Sub GetFocus(c As Control)
    c.SelStart = 0
    c.SelLength = Len(c.Text)
End Sub


Private Sub txt_String_Validate(Cancel As Boolean)
   txt_String = Left(txt_String & "    ", 4)
End Sub
