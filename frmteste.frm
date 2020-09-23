VERSION 5.00
Object = "*\Aprjmciconmenu.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmteste 
   Caption         =   "Test of MCunha98 - Icons in Menus"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin prjmciconmenu.mciconmenu mciconmenu1 
      Left            =   1200
      Top             =   3360
      _ExtentX        =   1058
      _ExtentY        =   1058
   End
   Begin VB.Frame Fundo 
      Caption         =   " Menus Appearance "
      Height          =   1215
      Index           =   3
      Left            =   3360
      TabIndex        =   16
      Top             =   2280
      Width           =   4815
      Begin VB.CommandButton CmdAddMenu 
         Caption         =   "Add new menus to Help Menu"
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox ChkBoldOpen 
         Caption         =   "Set menu OPEN to default (bold)"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Fundo 
      Caption         =   " Change the background picture "
      Height          =   2175
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton CmdChangeBackGround 
         Caption         =   "Change background..."
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   4575
      End
      Begin VB.PictureBox picBackGround 
         Height          =   1335
         Left            =   120
         Picture         =   "frmteste.frx":0000
         ScaleHeight     =   1275
         ScaleWidth      =   4515
         TabIndex        =   14
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Basici functions of control "
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   3135
      Begin VB.CheckBox ChkUseBackground 
         Caption         =   "Use background in menus"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox ChkEnableCopy 
         Caption         =   "Enable/Disable ""Copy"""
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox ChkHighLight 
         Caption         =   "HighLight Menus in Button Style"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Fundo 
      Caption         =   " Changing caption of Menu in run-time "
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
      Begin VB.TextBox TxtCaption 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton CmdChangeCaption 
         Caption         =   "Change !"
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         Caption         =   "Change the caption of menu EXIT:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.Frame Fundo 
      Caption         =   " Changing icons in run-time "
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
      Begin VB.PictureBox picSearch 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton CmdChangeSearchIcon 
         Caption         =   "Change icon"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         Caption         =   "The current icon of menu SEARCH is:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2685
      End
   End
   Begin MSComDlg.CommonDialog CDial 
      Left            =   2760
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":0C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":0D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":0EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":10DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":1236
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":1392
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":192E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":1A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":1BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":1D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":1E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":1FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":2156
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":22B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":240E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":25EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":2746
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":28A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":2E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":33DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmteste.frx":3536
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMenuVB 
      Height          =   495
      Left            =   120
      Picture         =   "frmteste.frx":3AD2
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   18
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu MnuArquivo 
      Caption         =   "&File"
      Begin VB.Menu MnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu MnuSendTo 
         Caption         =   "&Send to..."
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuSetPrinter 
         Caption         =   "&Set Printer"
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "Re&do"
      End
      Begin VB.Menu sp4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu MnuDel 
         Caption         =   "&Delete"
      End
      Begin VB.Menu sp6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSelectAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu MnuFind 
      Caption         =   "&Find"
      Begin VB.Menu MnuSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu MnuSearchNext 
         Caption         =   "Search next..."
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuTopics 
         Caption         =   "&Topics..."
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu MnuWebSite 
         Caption         =   "&Visit MCunha98 WebSite"
      End
   End
End
Attribute VB_Name = "frmteste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------
'By Mauricio Cunha
'www.go.to/mcunha98
'------------------------------------------------------
Public TxtHelp As String
Public myLONG As Long
Public myABOUT As Long
Public myFile As String

Private Sub ChkBoldOpen_Click()
mciconmenu1.MenuDefault("MnuOpen") = CBool(ChkBoldOpen.Value)
End Sub

Private Sub ChkEnableCopy_Click()
Me.mciconmenu1.Enabled("MnuCopy") = CBool(ChkEnableCopy.Value)
End Sub

Private Sub ChkHighLight_Click()
mciconmenu1.HighlightStyle = ChkHighLight.Value
End Sub


Private Sub ChkUseBackground_Click()
 If CBool(ChkUseBackground.Value) = True Then
  Set mciconmenu1.BackgroundPicture = picBackGround
 Else
  mciconmenu1.ClearBackgroundPicture
 End If
End Sub

Private Sub CmdAddMenu_Click()
If mciconmenu1.MenuExists("MnuVB") = True Or mciconmenu1.MenuExists("MnuVBSep") = True Then
 MsgBox "The menu 'MnuVB' is already loaded !!!", 16
 CmdAddMenu.Enabled = False
 Exit Sub
Else
 myLONG = mciconmenu1.MenuIndex("MnuHelp")
 mciconmenu1.AddItem "-", "MnuVBSep", "This menu is added by code", , myLONG
 mciconmenu1.AddItem "Add by Code - Visual Basic", "MnuVB", "New menu add", , myLONG
 ChkLogoVBVisible.Enabled = True
End If

End Sub

Private Sub CmdChangeBackGround_Click()
On Error Resume Next

myFile = ""
CDial.FileName = ""
CDial.DialogTitle = "Select one picture..."
CDial.Filter = "Image files (*.jpg,*.bmp,*.gif)|*.jpg;*.bmp"
CDial.ShowOpen

If Trim(CDial.FileName = "") Then
 picBackGround.Cls
 ChkUseBackground.Value = 0
Else
 picBackGround.Picture = LoadPicture(CDial.FileName)
 Set mciconmenu1.BackgroundPicture = picBackGround
End If

End Sub

Private Sub CmdChangeCaption_Click()
mciconmenu1.Caption("MnuExit") = TxtCaption.Text
End Sub


Private Sub CmdChangeSearchIcon_Click()
If mciconmenu1.ItemIcon("MnuSearch") = 15 Then
 mciconmenu1.ItemIcon("MnuSearch") = ImageList1.ListImages.Item(18).Index - 1
 picSearch.Picture = ImageList1.ListImages.Item(18).Picture
Else
 mciconmenu1.ItemIcon("MnuSearch") = ImageList1.ListImages.Item(16).Index - 1
 picSearch.Picture = ImageList1.ListImages.Item(16).Picture
End If
End Sub


Private Sub Form_Load()
mciconmenu1.SubClassMenu Me
mciconmenu1.ImageList = ImageList1
Set mciconmenu1.BackgroundPicture = picBackGround

mciconmenu1.ItemIcon("MnuNew") = ImageList1.ListImages.Item(1).Index - 1
mciconmenu1.ItemIcon("MnuOpen") = ImageList1.ListImages.Item(2).Index - 1
mciconmenu1.ItemIcon("MnuSave") = ImageList1.ListImages.Item(3).Index - 1
mciconmenu1.ItemIcon("MnuSaveAs") = ImageList1.ListImages.Item(4).Index - 1
'------------------------------------
mciconmenu1.ItemIcon("MnuClose") = ImageList1.ListImages.Item(5).Index - 1
mciconmenu1.ItemIcon("MnuSendTo") = ImageList1.ListImages.Item(6).Index - 1
'------------------------------------
mciconmenu1.ItemIcon("MnuPrint") = ImageList1.ListImages.Item(7).Index - 1
mciconmenu1.ItemIcon("MnuSetPrinter") = ImageList1.ListImages.Item(8).Index - 1
'------------------------------------
mciconmenu1.ItemIcon("MnuExit") = ImageList1.ListImages.Item(9).Index - 1
'------------------------------------
mciconmenu1.ItemIcon("MnuUndo") = ImageList1.ListImages.Item(10).Index - 1
mciconmenu1.ItemIcon("MnuRedo") = ImageList1.ListImages.Item(11).Index - 1
'------------------------------------------------------
mciconmenu1.ItemIcon("MnuCut") = ImageList1.ListImages.Item(12).Index - 1
mciconmenu1.ItemIcon("MnuCopy") = ImageList1.ListImages.Item(13).Index - 1
mciconmenu1.ItemIcon("MnuPaste") = ImageList1.ListImages.Item(14).Index - 1
mciconmenu1.ItemIcon("MnuDel") = ImageList1.ListImages.Item(15).Index - 1
'------------------------------------------------------
mciconmenu1.ItemIcon("MnuSearch") = ImageList1.ListImages.Item(16).Index - 1
mciconmenu1.ItemIcon("MnuSearchNext") = ImageList1.ListImages.Item(17).Index - 1
picSearch.Picture = ImageList1.ListImages.Item(16).ExtractIcon


mciconmenu1.ItemIcon("MnuTopics") = ImageList1.ListImages.Item(19).Index - 1
mciconmenu1.ItemIcon("MnuAbout") = ImageList1.ListImages.Item(20).Index - 1
mciconmenu1.ItemIcon("MnuWebSite") = ImageList1.ListImages.Item(21).Index - 1

'------------------------------------------------------
'Add the microsoft visual basic logo to help menu
'------------------------------------------------------
'mciconmenu1.ItemPicture 2, 27, picMenuVB.Picture
'------------------------------------------------------
'Aling the help menu in right
'------------------------------------------------------
mciconmenu1.ItemRight "Help", 3 '<-File + Edit + Search
'------------------------------------------------------
'Remove X button
'------------------------------------------------------
myLONG = mciconmenu1.SystemMenuCount
mciconmenu1.SystemMenuRemoveItem myLONG
'------------------------------------------------------
'Get handle of menu About (see the system menu for more details...
'------------------------------------------------------
myABOUT = mciconmenu1.SystemMenuAppendItem("&About this control")
End Sub


Private Sub mciconmenu1_InitPopupMenu(ParentItemNumber As Long)
Select Case ParentItemNumber
 Case 1: TxtHelp = "As opções deste menu dão acessibilidade à funções de gerenciamento de arquivos"
 Case 14: TxtHelp = "As opções deste menu são relativas as edições possíveis dentro do arquivo"
 Case 24: TxtHelp = "Opções de busca simples e avançada dentro do arquivo..."
 Case Else: TxtHelp = ""
End Select
SBar.SimpleText = "(" & ParentItemNumber & ") " & TxtHelp
End Sub

Private Sub mciconmenu1_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)

Select Case ItemNumber
 'Menu arquivo...
 Case 1: TxtHelp = "Options of file menu"
 Case 2: TxtHelp = "Create a new text file"
 Case 3: TxtHelp = "Open one file"
 Case 4: TxtHelp = "Save the current file"
 Case 5: TxtHelp = "Save as the current file"
 '-----------------------------------------
 Case 7: TxtHelp = "Close this file"
 Case 8: TxtHelp = "Send this file to user in email"
 '-----------------------------------------
 Case 10: TxtHelp = "Print this file"
 Case 11: TxtHelp = "Set the printer of windows..."
 '-----------------------------------------
 Case 13: TxtHelp = "Close this programm"
 
 '-----------------------------------------
 Case 15: TxtHelp = "Undo actions..."
 Case 16: TxtHelp = "Redo actions..."
 Case 18: TxtHelp = "Cut the selected text..."
 Case 19: TxtHelp = "Copy the selected text..."
 Case 20: TxtHelp = "Get current content of clipboard area..."
 Case 21: TxtHelp = "Delete selected text..."
 Case 23: TxtHelp = "Select all text"
 
 Case 25: TxtHelp = "Find a keyword in text"
 Case 26: TxtHelp = "Find next ocorrency of keyword..."
 
 Case Else: TxtHelp = ""
End Select

SBar.SimpleText = "(" & ItemNumber & ") " & TxtHelp
End Sub

Private Sub mciconmenu1_SystemMenuClick(ItemNumber As Long)
If ItemNumber = myABOUT Then
 mciconmenu1.AboutBox
End If
End Sub

Private Sub mciconmenu1_SystemMenuItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
SBar.SimpleText = "(" & ItemNumber & ")"
 If ItemNumber = myABOUT Then
  SBar.SimpleText = "Show the aboutbox message..."
 End If
End Sub

Private Sub MnuExit_Click()
Unload Me
End Sub

Private Sub MnuWebSite_Click()
mciconmenu1.OpenURL "http://www.go.to/mcunha98"
End Sub
