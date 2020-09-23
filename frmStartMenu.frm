VERSION 5.00
Begin VB.Form frmProgramsMenu 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2880
      Top             =   540
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   975
      TabIndex        =   5
      Top             =   1005
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1035
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picBlackArrow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1125
      Picture         =   "frmStartMenu.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picWhiteArrow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      Picture         =   "frmStartMenu.frx":036D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   615
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2235
      TabIndex        =   0
      Top             =   1065
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   300
      Index           =   0
      Left            =   0
      Picture         =   "frmStartMenu.frx":06DA
      ScaleHeight     =   300
      ScaleWidth      =   4035
      TabIndex        =   4
      Top             =   0
      Width           =   4035
   End
End
Attribute VB_Name = "frmProgramsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a remixed code from an already remixed code I think
'the already modified code comes from the BoS Project. So Just
'want to give credits to the programmer(s) who created this code
'in the first hand.

Option Explicit

Private OldIndex As Integer
Private MaxLen As Long 'Integer

Private fChild As frmProgramsMenu
Private mParent As frmProgramsMenu
Private CurY As Long
Private CurYIcon As Long

Private Sub Form_Activate()
  frmProgramsMenu.Top = (Screen.Height - frmProgramsMenu.Height) / 2
  frmProgramsMenu.Left = 590
End Sub

Private Sub Form_Load()
  Me.AutoRedraw = False 'faster
  Me.ScaleMode = 3 'pixel
  OldIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (fChild Is Nothing) Then Unload fChild
End Sub

Public Sub UnloadAll()
  If Not (mParent Is Nothing) Then mParent.UnloadAll
  Unload Me
End Sub


Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 If Button = vbLeftButton Then
    If Index > Dir1.ListCount - 1 Then
        ShellExecute Me.hWnd, "open", Dir1.path & "\" & Right(picItem(Index).Tag, Len(picItem(Index).Tag) - 8), "", "", 1
        mParent.UnloadAll
        Unload Me
        'HideStartMenu
        's_Playsound "select"
    End If
  ElseIf Button = vbRightButton Then
  
  End If
End Sub

Private Sub picItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Dir1.ListCount + File1.ListCount = 0 Then Exit Sub 'if there is nothing to show
  If Index <> OldIndex Then 'Over(Index) = False Then
    If OldIndex = -1 Then OldIndex = 0
    
    ' --- reset the old selection ---
    picItem(OldIndex).Cls
    picItem(OldIndex).BackColor = vbBlack
    picItem(OldIndex).ForeColor = vbButtonFace
    picItem(OldIndex).CurrentY = CurY
    picItem(OldIndex).Print picItem(OldIndex).Tag
    
    If OldIndex >= Dir1.ListCount Then
      DrawIcon Dir1.path & "\" & File1.List(OldIndex - Dir1.ListCount), OldIndex
    Else
      BitBlt picItem(OldIndex).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
      DrawIcon Dir1.List(OldIndex), OldIndex
    End If
    
    ' --- highlite new selection ---
    picItem(Index).Cls
    picItem(Index).BackColor = vbBlack
    picItem(Index).ForeColor = vbWhite
    picItem(Index).CurrentY = CurY
    picItem(Index).Print picItem(Index).Tag
    If Index >= Dir1.ListCount Then  'If not a directory then
      DrawIcon Dir1.path & "\" & File1.List(Index - Dir1.ListCount), Index ', False
    Else 'index < Dir1.ListCount
      BitBlt picItem(Index).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 16, 16, picWhiteArrow.hdc, 0, 0, vbSrcCopy
      DrawIcon Dir1.List(Index), Index ', False
      
      ' --- show new child menu ---
      Timer1.Interval = 100
    End If
    picTemp.Line (0, 0)-(19, 0)
    picTemp.Line (0, 0)-(0, 19)
    picTemp.ForeColor = vbBlack
    picTemp.Line (0, 19)-(19, 19)
    picTemp.Line (19, 0)-(19, 19)
    BitBlt picItem(Index).hdc, 0, CurYIcon, 20, 20, picTemp.hdc, 0, 0, vbSrcCopy
    
    
    OldIndex = Index
    If Not (fChild Is Nothing) Then Unload fChild
    
    If Index < Dir1.ListCount Then
'          s_Playsound "open"
    Else
'          s_Playsound "hover"
    End If
  End If
End Sub

Public Sub GetMenu(path As String, Optional Parent As frmProgramsMenu = Nothing)
  Dim i As Long
  Dim lTemp As Long
  
  Set mParent = Parent
  MaxLen = 0
  
  Dir1.path = path
  File1.path = path
  If File1.ListCount + Dir1.ListCount = 0 Then
      picItem(0).CurrentY = ((picItem(0).Height * Screen.TwipsPerPixelY) - picItem(0).TextHeight("|")) / 2
      picItem(0).Print "[ Empty ]"
      MaxLen = picItem(0).TextWidth("[ Empty ]")
  Else
      For i = 1 To Dir1.ListCount + File1.ListCount - 1
          Load picItem(i)
          picItem(i).Visible = True
          picItem(i).Top = picItem(0).Height * i
      Next
        CurYIcon = ((picItem(0).Height) - 20) / 2
        CurY = ((picItem(0).Height * Screen.TwipsPerPixelY) - picItem(0).TextHeight("|")) / 2
        For i = 0 To Dir1.ListCount - 1
            DrawIcon Dir1.List(i), i
            picItem(i).CurrentY = CurY
            picItem(i).Tag = "        " & ExtractFileName(Dir1.List(i))
            picItem(i).Print picItem(i).Tag
            lTemp = picItem(i).TextWidth(picItem(i).Tag)
            If lTemp > MaxLen Then MaxLen = lTemp
        Next
        For i = 0 To File1.ListCount - 1
            picTemp.BackColor = vbBlack
            DrawIcon Dir1.path & "\" & File1.List(i), i + Dir1.ListCount
            picItem(i + Dir1.ListCount).CurrentY = CurY
            picItem(i + Dir1.ListCount).Tag = "        " & Left(File1.List(i), Len(File1.List(i)) - 4)
            picItem(i + Dir1.ListCount).Print picItem(i + Dir1.ListCount).Tag
            lTemp = picItem(i + Dir1.ListCount).TextWidth(picItem(i + Dir1.ListCount).Tag)
            If lTemp > MaxLen Then MaxLen = lTemp
        Next
  End If
  Me.Width = MaxLen + 500
  Me.Height = (picItem.Count * picItem(0).Height * Screen.TwipsPerPixelY) - 150
  If mParent Is Nothing Then
    Me.Height = Me.Height
  End If
  
  If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
  SetWindowPos Me.hWnd, -1, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.ScaleWidth, Me.ScaleHeight + 10, SWP_NOREPOSITION
  For i = 0 To picItem.Count - 1
      picItem(i).Width = Me.ScaleWidth
  Next
  For i = 0 To Dir1.ListCount - 1
      BitBlt picItem(i).hdc, Me.ScaleWidth - 18, CurY / Screen.TwipsPerPixelY, 18, 18, picBlackArrow.hdc, 0, 0, vbSrcCopy
  Next
  
 
  Me.Show
  Me.Refresh
  
End Sub

Sub DrawIcon(path, Index, Optional blt = True)
  Dim hImgLarge&
  
  hImgLarge& = SHGetFileInfo(path, 0&, shinfo, Len(shinfo), _
  BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
  picTemp.Cls
  If blt Then
      ImageList_Draw hImgLarge&, shinfo.iIcon, picTemp.hdc, 2, 2, ILD_TRANSPARENT
      BitBlt picItem(Index).hdc, 0, CurYIcon, 20, 20, picTemp.hdc, 0, 0, vbSrcCopy
  Else
      ImageList_Draw hImgLarge&, shinfo.iIcon, picTemp.hdc, 2, 2, ILD_TRANSPARENT
  End If
End Sub

Private Sub Timer1_Timer()
  Timer1.Interval = 0
  
  ' --- show new child menu ---
  If OldIndex < Dir1.ListCount Then
    Set fChild = New frmProgramsMenu
    fChild.Top = Me.Top + picItem(OldIndex).Top * Screen.TwipsPerPixelX
    fChild.Left = Me.Left + Me.Width + 50
    fChild.GetMenu Dir1.path & "\" & Right(picItem(OldIndex).Tag, Len(picItem(OldIndex).Tag) - 8), Me
  End If
End Sub


