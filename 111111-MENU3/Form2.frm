VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "Menu Engine/Compiler V0.99 by Vanja Fuckar,EMAIL:INGA@VIP.HR"
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   7890
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3120
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000080FF&
      Caption         =   "Run (Another Menues Object) Owner Draw Sample"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   1680
      Picture         =   "Form2.frx":030A
      ScaleHeight     =   900
      ScaleWidth      =   1125
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Load Owner Draw Sample 5"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Save Binary File"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Load Binary File"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Load Sample 4"
      Height          =   375
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Load Sample 3"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Load Sample 2"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Load Sample 1"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00F1D2C9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   240
      Width           =   6375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000C000&
      Caption         =   "Track Pop Up"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Load Text File"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Save Text File"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Compile And Menu Bar"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Menu Compiler Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   5175
   End
   Begin VB.Menu TemporaryRequired 
      Caption         =   ""
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MM As New MenuMaker
Private Mhandle As Long
Private WithEvents MENU1 As Menues
Attribute MENU1.VB_VarHelpID = -1
Private WithEvents MENU2 As Menues
Attribute MENU2.VB_VarHelpID = -1
Dim sFile As String
Dim sPath As String
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020


Const ODS_SELECTED = &H1
Const ODS_GRAYED = &H2
Const ODS_DISABLED = &H4
Const ODS_CHECKED = &H8
Const ODS_FOCUS = &H10
Const ODS_DEFAULT = &H20
Const ODS_HOTLIGHT = &H40
Const ODS_INACTIVE = &H80
Const ODS_NOACCEL = &H100
Const ODS_NOFOCUSRECT = &H200

Const DT_TOP = &H0
 Const DT_LEFT = &H0
Const DT_CENTER = &H1
 Const DT_RIGHT = &H2
 Const DT_VCENTER = &H4
Const DT_BOTTOM = &H8
Const DT_WORDBREAK = &H10
 Const DT_SINGLELINE = &H20
Const DT_EXPANDTABS = &H40
 Const DT_TABSTOP = &H80
 Const DT_NOCLIP = &H100
 Const DT_EXTERNALLEADING = &H200
Const DT_CALCRECT = &H400
Const DT_NOPREFIX = &H800
Const DT_INTERNAL = &H1000
Const DT_EDITCONTROL = &H2000
Const DT_PATH_ELLIPSIS = &H4000
Const DT_END_ELLIPSIS = &H8000
Const DT_MODIFYSTRING = &H10000
Const DT_RTLREADING = &H20000
Const DT_WORD_ELLIPSIS = &H40000





Private Sub Command1_Click()
MM.MenuData = Text1
Mhandle = MM.CompileMenu
MENU1.ImportMenu Mhandle
MENU1.ShowMenu Me.hwnd
End Sub

Private Sub Command10_Click()
If MM.Handle = 0 Then MsgBox "Compile First!", vbInformation, "Info": Exit Sub
Dim aa As Boolean
aa = GetSaveFilePath(hwnd, "", 0, "", "", "", "Save As Text File", sPath)
If aa = False Then Exit Sub
MM.SaveBinaryMenu sPath
End Sub

Private Sub Command11_Click()
Text1 = LoadResString(505)
End Sub

Private Sub Command12_Click()
MsgBox "Presentation another Menues Object with same Form!" & vbCrLf & _
"Usually DO NOT USE MORE THAN ONE OBJECT PER FORM!!! BECAUSE OF SUBCLASSING LIMITATION!", vbExclamation, "WARNING!"

Text1 = LoadResString(506)
Text1 = Text1
MM.MenuData = Text1
Text1 = Text1 & vbCrLf & "Don not Compile That!!! " & vbCrLf & "Because Menu allready running in another OBJECT!"

Mhandle = MM.CompileMenu
MENU2.ImportMenu Mhandle
MENU2.ShowMenu Me.hwnd
MENU2.TrackMenu 0, 100, 100, Me.hwnd
End Sub

Private Sub Command2_Click()
If Text1 = "" Then Exit Sub
Dim aa As Boolean
aa = GetSaveFilePath(hwnd, "", 0, "", "", "", "Save As Binary File", sPath)
If aa = False Then Exit Sub
Dim FFL As Long
FFL = FreeFile
If Dir(sPath) <> "" Then Kill sPath
Open sPath For Binary As #FFL
Put #1, , Text1.Text
Close #1
End Sub

Private Sub Command3_Click()
Text1 = LoadResString(503)
End Sub

Private Sub Command4_Click()
Dim aa As Boolean
aa = GetOpenFilePath(hwnd, "", 0, sFile, "", "Load Text File", sPath)
If aa = False Then Exit Sub
Dim FFL As Long
FFL = FreeFile
Dim STRX As String
Open sPath For Binary Access Read As #FFL
STRX = Space(LOF(FFL))
Get #FFL, , STRX
Close #1
Text1 = STRX
STRX = ""
End Sub



Private Sub Command5_Click()
Text1 = LoadResString(504)
End Sub

Private Sub Command6_Click()
Text1 = LoadResString(501)
End Sub

Private Sub Command7_Click()

MENU1.TrackMenu 0, 100, 200, Me.hwnd

End Sub

Private Sub Command8_Click()
Text1 = LoadResString(502)
End Sub

Private Sub Command9_Click()
Dim aa As Boolean
aa = GetOpenFilePath(hwnd, "", 0, sFile, "", "Load Binary File", sPath)
If aa = False Then Exit Sub
MENU1.LoadBinaryMenu sPath
MENU1.ShowMenu hwnd
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Text2 = LoadResString(500)
Set MENU1 = New Menues
Set MENU2 = New Menues
SND1 = LoadResData(101, "CUSTOM")
SND2 = LoadResData(102, "CUSTOM")
End Sub




Private Sub MENU1_DrawItem(drawIstruct As MenuObject.DRAWITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
StretchBlt drawIstruct.hdc, 0, _
0, 95, _
(drawIstruct.rcItem.Bottom) * _
5, Picture1.hdc, 0, 0, CLng(Picture1.Width / 15), _
CLng(Picture1.Height / 15), SRCCOPY

End Sub

Private Sub MENU1_MeasureItem(measureIstruct As MenuObject.MEASUREITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
measureIstruct.itemWidth = 79
measureIstruct.ItemHeight = 17
End Sub

Private Sub MENU1_MenuClick(ByVal ID As Long)
MsgBox "Menu Click ID:" & ID, vbInformation, "MENU OBJECT 1"
End Sub



Private Sub MENU2_DrawItem(drawIstruct As MenuObject.DRAWITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
Dim FNT1 As Long
Dim OO As Long
SetBkMode drawIstruct.hdc, 1
Dim Sbr As Long

Dim STTate As Long
Dim RCTX As RECT
RCTX = drawIstruct.rcItem

If (drawIstruct.itemState = ODS_SELECTED) Or (drawIstruct.itemState = 257) Then
Call DrawEdge(drawIstruct.hdc, RCTX, EDGE_RAISED, BF_RECT)
Call InflateRect(RCTX, -2, -2)
Sbr = CreateSolidBrush(&HFF&)
FillRect drawIstruct.hdc, RCTX, Sbr
DeleteObject Sbr

StretchBlt drawIstruct.hdc, RCTX.Left, _
RCTX.Top, 32, _
28, Picture2.hdc, 0, 0, CLng(Picture2.Width / 15), _
CLng(Picture2.Height / 15), SRCCOPY


If drawIstruct.itemID > 2 Then
PlaySound_Res ByVal VarPtr(SND1(0)), 0, &H4 Or &H1
Else
PlaySound_Res ByVal VarPtr(SND2(0)), 0, &H4 Or &H1
End If

Else
Sbr = CreateSolidBrush(&H8822&)
FillRect drawIstruct.hdc, RCTX, Sbr
DeleteObject Sbr
End If
SetTextColor drawIstruct.hdc, &HFFFF&
fnt = GetFont("Arial Narrow", 14, 0, False, False, False, FW_BOLD)
OO = SelectObject(drawIstruct.hdc, fnt)
Dim DTP As DRAWTEXTPARAMS
DTP.cbSize = Len(DTP)

If MenuText = "" Then MenuText = "Item " & drawIstruct.itemID 'WIN98 doesnt support ItemTextPointer.--->

Call DrawTextEx(drawIstruct.hdc, MenuText, Len(MenuText), drawIstruct.rcItem, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, DTP)
Call SelectObject(drawIstruct.hdc, OO)
DeleteObject FNT1
End Sub

Private Sub MENU2_MeasureItem(measureIstruct As MenuObject.MEASUREITEMSTRUCT, ByVal MenuText As String, ByVal ParentHwnd As Long)
measureIstruct.itemWidth = 300
measureIstruct.ItemHeight = 32
End Sub

Private Sub MENU2_MenuClick(ByVal ID As Long)
MsgBox "Second MENUES Object,Click Menu ID:" & ID, vbOKOnly, "MENU OBJECT 2"
End Sub
