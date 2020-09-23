VERSION 5.00
Object = "{570928AD-1209-11D3-967B-B4129805661E}#5.0#0"; "CSTRAY.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3765
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "FORM2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3045
      Left            =   90
      TabIndex        =   4
      Top             =   615
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   5371
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PID"
         Object.Width           =   1288
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Parent"
         Object.Width           =   1288
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   90
      ScaleHeight     =   1155
      ScaleWidth      =   4725
      TabIndex        =   10
      Top             =   2190
      Width           =   4755
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2205
      Top             =   4020
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   90
      ScaleHeight     =   1155
      ScaleWidth      =   4725
      TabIndex        =   9
      Top             =   615
      Width           =   4755
   End
   Begin csTrayOCX.csTray csTray1 
      Left            =   4230
      Top             =   435
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "FORM2.frx":08CA
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   135
      TabIndex        =   18
      Top             =   3405
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   150
      TabIndex        =   17
      Top             =   3420
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   1830
      Width           =   780
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   135
      TabIndex        =   15
      Top             =   1845
      Width           =   630
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kernel :: Threads"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2910
      TabIndex        =   14
      Top             =   3390
      Width           =   1905
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kernel :: Threads"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2925
      TabIndex        =   13
      Top             =   3405
      Width           =   1905
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CPU :: Usage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2910
      TabIndex        =   12
      Top             =   1815
      Width           =   1905
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CPU :: Usage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2925
      TabIndex        =   11
      Top             =   1830
      Width           =   1905
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stats."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDA04F&
      Height          =   240
      Left            =   1755
      TabIndex        =   8
      ToolTipText     =   "Show local computer informations."
      Top             =   375
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   255
      TabIndex        =   7
      ToolTipText     =   "Show all runnings process."
      Top             =   375
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stats."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1770
      TabIndex        =   6
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   5
      Top             =   390
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   1575
      Picture         =   "FORM2.frx":11A4
      Top             =   330
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   105
      Picture         =   "FORM2.frx":1CA6
      Top             =   330
      Width           =   1440
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00755433&
      X1              =   -90
      X2              =   4950
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   4920
      X2              =   4920
      Y1              =   0
      Y2              =   3900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4935
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4680
      TabIndex        =   1
      Top             =   30
      Width           =   195
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4440
      TabIndex        =   0
      Top             =   30
      Width           =   210
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00755433&
      X1              =   4845
      X2              =   4845
      Y1              =   60
      Y2              =   210
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00755433&
      X1              =   4695
      X2              =   4860
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00755433&
      X1              =   4620
      X2              =   4620
      Y1              =   45
      Y2              =   195
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   4470
      X2              =   4635
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   4710
      Picture         =   "FORM2.frx":27A8
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   4485
      Picture         =   "FORM2.frx":28E6
      Top             =   60
      Width           =   135
   End
   Begin VB.Shape Shape4 
      Height          =   195
      Left            =   4680
      Top             =   30
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape6 
      Height          =   195
      Left            =   4455
      Top             =   30
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   4470
      Top             =   45
      Width           =   165
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   4695
      Top             =   45
      Width           =   165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Manipulator - Process Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   60
      TabIndex        =   2
      Top             =   15
      Width           =   4365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manipulator - Process Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   30
      Width           =   4365
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99063&
      FillColor       =   &H009B6F43&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   -915
      Picture         =   "FORM2.frx":2A24
      Top             =   15
      Width           =   5850
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   3510
      Left            =   0
      Top             =   255
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const REG_DWORD = 4
Private Const HKEY_DYN_DATA = &H80000006

Private Type LARGE_INTEGER
 lowpart As Long
 highpart As Long
End Type

Dim Status, X_Initial, Y_Initial, Dist_Am
Dim Panel, Proc As Integer
Dim CpuUse(40), CpuCounter As Integer
Dim NbThreads(40), ThreadsCounter As Integer
Dim pid(200) As Long

Private Sub csTray1_MouseUp(Button As Integer)
 csTray1.Visible = False
 Me.WindowState = 0
 Me.Visible = True
End Sub

Private Sub Form_Load()
 Panel = 0
 ListView1.Visible = True
 CpuCounter = 0
 
 Call InitDraw(Picture1)
 Call InitDraw(Picture2)
 Call LoadTaskList(1)
 Call InitCPU
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 1
 X_Initial = X
 Y_Initial = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + X - X_Initial
  Me.Top = Me.Top + Y - Y_Initial
 Else
  Call remap
 End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 0
 Dist_Am = 100
 
 If Me.Left < Dist_Am Then Me.Left = 0
 If Me.Top < Dist_Am Then Me.Top = 0
 If Me.Left + Me.Width > Screen.Width - Dist_Am Then Me.Left = Screen.Width - Me.Width
 If Me.Top + Me.Height > Screen.Height - Dist_Am Then Me.Top = Screen.Height - Me.Height
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value
 End
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape4.Visible = False Then
  Call remap
  Shape4.Visible = True
 End If
End Sub

Sub remap()
 If Shape4.Visible = True Then Shape4.Visible = False
 If Shape6.Visible = True Then Shape6.Visible = False
End Sub

Private Sub Label5_Click()
 If Panel <> 0 Then
  Panel = 0
  Label5.ForeColor = &HFFFFFF
  Label6.ForeColor = &HFDA04F
  ListView1.Visible = True
 End If
End Sub

Private Sub Label6_Click()
 If Panel <> 1 Then
  Panel = 1
  Label6.ForeColor = &HFFFFFF
  Label5.ForeColor = &HFDA04F
  ListView1.Visible = False
 End If
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line2.BorderColor
 Line2.BorderColor = Shape5.BorderColor
 Line3.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape6.Visible = False Then
  Call remap
  Shape6.Visible = True
 End If
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line2.BorderColor
 Line2.BorderColor = Shape5.BorderColor
 Line3.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
 
 Me.WindowState = 1
 Me.Visible = False
 csTray1.Visible = True
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Function LoadTaskList(mode) As Boolean
 Dim CurrWnd As Long
 Dim Length As Long
 Dim TaskName As String
 Dim Parent As Long

 CurrWnd = GetWindow(Me.hWnd, GW_HWNDFIRST)

 If mode = 1 Then Form2.ListView1.ListItems.Clear: Proc = 0
 noProc = 0
 While CurrWnd <> 0
  Parent = GetParent(CurrWnd)
  Length = GetWindowTextLength(CurrWnd)
  TaskName = Space$(Length + 1)
  Length = GetWindowText(CurrWnd, TaskName, Length + 1)
  TaskName = Left$(TaskName, Len(TaskName) - 1)

  If Length > 0 Then
   If TaskName <> Me.Caption Then
    If mode = 1 Then
     Proc = Proc + 1
     pid(Proc) = CurrWnd
     Set lvItem = Form2.ListView1.ListItems.Add(, , TaskName)
     lvItem.SubItems(1) = Str$(CurrWnd)
     lvItem.SubItems(2) = Str$(Parent)
    Else
     noProc = noProc + 1
     Found = False
     For i = 1 To Proc
      If pid(i) = CurrWnd Then Found = True
     Next i
     If Found = False Then LoadTaskList = False: Exit Function
    End If
   End If
  End If
  CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
  DoEvents
 Wend
 If noProc = Proc Then
  LoadTaskList = True
 Else
  LoadTaskList = False
 End If
End Function

Private Sub InitDraw(PicHandle As PictureBox)
 PicHandle.Cls
 PicHandle.Line (0, PicHandle.Height / 2)-(PicHandle.Width, PicHandle.Height / 2), RGB(0, 0, 0)
 PicHandle.Line (0, PicHandle.Height / 4)-(PicHandle.Width, PicHandle.Height / 4), RGB(128, 128, 128)
 PicHandle.Line (0, PicHandle.Height / 2 + PicHandle.Height / 4)-(PicHandle.Width, PicHandle.Height / 2 + PicHandle.Height / 4), RGB(128, 128, 128)
 For i = 0 To PicHandle.Width Step PicHandle.Width / 40
  PicHandle.Line (i, 0)-(i, PicHandle.Height), RGB(128, 128, 128)
 Next i
End Sub

Private Sub InitCPU()
 Dim lData As Long
 Dim lType As Long
 Dim lSize As Long
 Dim hKey As Long
 Dim Qry As String
    
 Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
 If Qry <> 0 Then MsgBox "Could not open registery!": End
 lType = REG_DWORD
 lSize = 4
 Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
 Qry = RegQueryValueEx(hKey, "KERNEL\Threads", 0, lType, lData, lSize)
 Qry = RegCloseKey(hKey)
End Sub

Private Sub Timer1_Timer()
 Dim lData As Long
 Dim lType As Long
 Dim lSize As Long
 Dim hKey As Long
 Dim Qry As String
 Dim Status As Long
                  
 Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
                
 If Qry <> 0 Then MsgBox "Could not open registery!": End
                
 lType = REG_DWORD
 lSize = 4
 Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
 CPU = Int(lData)
 Qry = RegQueryValueEx(hKey, "KERNEL\Threads", 0, lType, lData, lSize)
 Threads = Int(lData)
 Qry = RegCloseKey(hKey)

 If CpuCounter < 40 Then
  CpuCounter = CpuCounter + 1
  CpuUse(CpuCounter) = CPU
 Else
  For i = 1 To 40
   CpuUse(i - 1) = CpuUse(i)
  Next i
  CpuUse(CpuCounter) = CPU
 End If
 Call InitDraw(Picture1)
 yScale = Picture1.Height / 100
 xScale = Picture1.Width / 40
 For i = 0 To 39
  Y1 = Picture1.Height - (CpuUse(i) * yScale)
  X1 = i * xScale
  Y2 = Picture1.Height - (CpuUse(i + 1) * yScale)
  X2 = (i + 1) * xScale
  Picture1.Line (X1, Y1)-(X2, Y2), RGB(60, 60, 200)
 Next i
 a$ = Right$(Str$(CPU), Len(Str$(CPU)) - 1) + " %"
 If a$ <> Label13.Caption Then
  Label13.Caption = a$
  Label14.Caption = a$
 End If
 
 If ThreadsCounter < 40 Then
  ThreadsCounter = ThreadsCounter + 1
  NbThreads(ThreadsCounter) = Threads
 Else
  For i = 1 To 40
   NbThreads(i - 1) = NbThreads(i)
  Next i
  NbThreads(ThreadsCounter) = Threads
 End If
 Call InitDraw(Picture2)
 yScale = Picture2.Height / 100
 xScale = Picture2.Width / 40
 For i = 0 To 39
  Y1 = Picture2.Height - (NbThreads(i) * yScale)
  X1 = i * xScale
  Y2 = Picture2.Height - (NbThreads(i + 1) * yScale)
  X2 = (i + 1) * xScale
  Picture2.Line (X1, Y1)-(X2, Y2), RGB(60, 60, 200)
 Next i
 a$ = Right$(Str$(Threads), Len(Str$(Threads)) - 1)
 If a$ <> Label15.Caption Then
  Label15.Caption = a$
  Label16.Caption = a$
 End If

 If LoadTaskList(0) = False Then LoadTaskList (1)
End Sub
