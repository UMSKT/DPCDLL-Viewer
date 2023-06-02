VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DPCDLL LicenseType Viewer"
   ClientHeight    =   8865
   ClientLeft      =   720
   ClientTop       =   1035
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Dialog.frx":0000
      Top             =   120
      Width           =   5895
   End
   Begin MSComctlLib.ListView List 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10186
      SortKey         =   4
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   847
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "-DD"
         Object.Width           =   847
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "min -BBB-"
         Object.Width           =   1552
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "max"
         Object.Width           =   847
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "LicType"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Activationdays"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "EvaluationDays"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8520
      Width           =   5775
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   8280
      Width           =   5775
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Option Explicit
Dim dpc As New FileStream
Private Sub Form_Load()
   On Error Resume Next
   Dim tmpbuffer$
   
   GetAttr "dpcdll.dll"
   If err Then
      tmpbuffer = Space(128)
      tmpbuffer = Left(tmpbuffer, GetSystemDirectory(tmpbuffer, 127))
   Else
      tmpbuffer = "."
   End If
   
  
   With dpc

 ' Open dpcdll.dll
Try:
      err.Clear
      .Create tmpbuffer & "\dpcdll.dll", , , True
     
     ' seek to last 50% of the file
      .Position = .Length / 2
     
     ' find a record with Actidays=30 and Eval=none - hope there is one
      .FindBytes &H1E, &H0, &H0, &H0, &HFF, &HFF, &HFF, &H7F
      If err <> 0 Then
         tmpbuffer = InputBox("Enter another path :)", "Opening dpcdll.dll failed", tmpbuffer)
         If tmpbuffer = "" Then End
         GoTo Try
      End If
     
     Label2.Caption = .FileName
     
     'Move back until to start (where 'index'==0)
      Do
         .Move -7 * 4
      Loop Until .longValue = 0
      
     'Fill listview with recorddata
      Dim index&, li As ListItem
      Do
         

         Set li = List.ListItems.Add(, , Dec2(index))
         
         Dim item
         For Each item In Array( _
            Dec2(.longValue \ 2), _
            Dec3(.longValue), _
            Dec3(.longValue), _
            Choose(.longValue, "Cooperate", "Retail", "Evaluation", "TablePC", "OEM", "Embedded"), _
            Filter_7fffffff(.longValue), _
            Filter_7fffffff(.longValue) _
            )
         
            li.ListSubItems.Add , , item
         Next
         
        'read Hashdatasize and skip following hashdata
         .Move .longValue
         index = index + 1
     'Do while index is increasing
      Loop While index = .longValue
      .CloseFile
   End With
   
   '
   Dim reg As New Registry
   reg.Create HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
   reg.RegValue = "ProductId"
   
   Dim PID$
   PID = reg.Regdata
   
   
   Dim BBB&, dd&
   dd = Left(Split(PID, "-")(3), 2)
   BBB = Split(PID, "-")(1)
   For Each li In List.ListItems
      If li.SubItems(1) = dd Then
         If (li.SubItems(2) <= BBB) And _
            (li.SubItems(3) >= BBB) Then
            li.Selected = True
            Label1.Caption = "Info: Line " & li & _
                              " matches to your current PID: " & PID
            
            Exit For
          End If
      End If
            
         
   Next
   
   BigNum.Show

End Sub

Private Function Filter_7fffffff$(value&)
   Filter_7fffffff = IIf(value = &H7FFFFFFF, "None", value)
End Function

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub List_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   List.Sorted = True
   If List.SortKey = ColumnHeader.index - 1 Then
      List.SortOrder = (List.SortOrder = lvwDescending) + 1
   Else
      List.SortKey = ColumnHeader.index - 1
   End If
End Sub

'Private Sub OKButton_Click()
'   Dim item
'   For Each item In List.ColumnHeaders
'   Debug.Print item.Width
'   Next
'End Sub
