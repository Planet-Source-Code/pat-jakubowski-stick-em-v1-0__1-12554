VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   3810
   ClientTop       =   3090
   ClientWidth     =   2340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   143
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stick'Em Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" _
    Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Const EM_GETLINECOUNT = &HBA
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_BOTTOM = 1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Sub cmdStick1_Click()

End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    txtNote.Text = GetSetting("StickyNote", "Info", "Text", "")
End Sub

Private Sub lblExit_Click()
    Dim Variable As VbMsgBoxResult
    Variable = MsgBox("Do you wish to save your Stick'em Note?", vbYesNoCancel, "Save?")
    If Variable = vbYes Then
        SaveSetting "StickyNote", "Info", "Text", txtNote.Text
        End
    ElseIf Variable = vbCancel Then
        Exit Sub
    Else
        SaveSetting "StickyNote", "Info", "Text", ""
        End
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        DragForm Me
    End If
End Sub

Private Sub txtNote_Change()
    Dim X As Integer
    X = 0
    Dim lCount As Long
    lCount = SendMessageLong(txtNote.hWnd, EM_GETLINECOUNT, 0, 0)
    Debug.Print "Lines: " & lCount
    For X = 6 To lCount
        txtNote.Height = lCount * 195
        frmMain.Height = txtNote.Height + 775
        cmdStick.Top = txtNote.Height + 500
        X = X + 1
    Next X
    lblTitle.Caption = txtNote.Text
End Sub

Private Sub cmdStick_Click()
    Me.Left = Screen.Width - Me.Width
    Me.Top = 0
    'SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    SetWindowPos frmMain.hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

