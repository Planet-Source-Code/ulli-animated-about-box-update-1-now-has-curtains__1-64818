VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   4575
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5460
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btExit 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   3345
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   3585
      Width           =   870
   End
   Begin VB.CommandButton btHold 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hold"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1290
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   3585
      Width           =   870
   End
   Begin VB.Timer tmrSlide 
      Interval        =   20
      Left            =   135
      Top             =   3870
   End
   Begin VB.PictureBox picViewport 
      BorderStyle     =   0  'Kein
      Height          =   3465
      Left            =   225
      ScaleHeight     =   231
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   334
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   210
      Width           =   5010
      Begin VB.PictureBox picCurtain 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000080&
         BorderStyle     =   0  'Kein
         Height          =   3330
         Index           =   1
         Left            =   0
         ScaleHeight     =   222
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   7
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   105
      End
      Begin VB.PictureBox picCurtain 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000080&
         BorderStyle     =   0  'Kein
         Height          =   3330
         Index           =   0
         Left            =   4770
         ScaleHeight     =   222
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   7
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   105
      End
      Begin VB.PictureBox picSlide 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Kein
         Height          =   3405
         Left            =   0
         ScaleHeight     =   227
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   308
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   4620
         Begin VB.PictureBox picIcon 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   0
            Left            =   2250
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   195
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lbVersion 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "Version"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   10
            Top             =   1290
            Visible         =   0   'False
            Width           =   4890
         End
         Begin VB.Label lbDivider 
            BackColor       =   &H00C0C0C0&
            Height          =   60
            Index           =   0
            Left            =   195
            TabIndex        =   8
            Top             =   3225
            Width           =   4620
         End
         Begin VB.Label lbOtherstuff2 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   " Otherstuff2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   705
            Index           =   0
            Left            =   75
            TabIndex        =   7
            Top             =   2445
            Visible         =   0   'False
            Width           =   4890
         End
         Begin VB.Label lbOtherstuff1 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "Otherstuff1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   6
            Top             =   2040
            Visible         =   0   'False
            Width           =   4890
         End
         Begin VB.Label lbTitle 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Index           =   0
            Left            =   75
            TabIndex        =   4
            Top             =   855
            Visible         =   0   'False
            Width           =   4890
         End
         Begin VB.Label lbCopyright 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   5
            Top             =   1650
            Visible         =   0   'False
            Width           =   4890
         End
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'animated about box
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'''''''
'HOW TO
'''''''
'
'    Load frmAbout
'    With frmAbout
'        .Theme = [Chose theme from 1..27]
'        .AppIcon([BackColor]) = Icon
'        .Title([ForeColor]) = "Testing the About Box"
'        .Version([ForeColor]) = "This is for the version"
'        .Copyright([ForeColor]) = "Enter your copyright here"
'        .Otherstuff1([ForeColor]) = "Enter other stuff here"
'        .Otherstuff2([ForeColor]) = "You may enter a longer desription of the project which spans several lines here"
'        .Show vbModal, Me
'    End With 'FRMABOUT
'
'    you can grab anywhere (exept the buttons) to move the window
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Enum MsgConsts
    HTCAPTION = 2
    WM_NCLBUTTONDOWN = 161
End Enum
#If False Then
Private HTCAPTION, WM_NCLBUTTONDOWN
#End If

Private Type Rect
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
Private WindowRect      As Rect

Private Type ColorTriple
    r As Long
    g As Long
    b As Long
End Type

Private CT              As ColorTriple

Private MouseIsDown     As Boolean
Private OriginalSpeed   As Long
Private myThemeColor    As Long
Private NormalDark      As Long
Private DarkerDark      As Long
Private hRgn1           As Long
Private hRgn2           As Long
Private Sidestep        As Long

Public Property Let AppIcon(BackColor As Long, nuIcon As Picture)

    picIcon(0).Visible = True
    picIcon(0).BackColor = BackColor
    picIcon(0).Picture = nuIcon
    Copy picIcon(1)

End Property

Private Sub btExit_Click()

    If Sidestep < 1 Then
        Sidestep = 1
        Do
            DoEvents
        Loop Until Sidestep = 0
        Sleep 800
        Unload Me
    End If

End Sub

Private Sub btExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    btExit.BackColor = DarkerDark

End Sub

Private Sub btHold_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseIsDown = True

End Sub

Private Sub btHold_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    btHold.BackColor = DarkerDark

End Sub

Private Sub btHold_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseIsDown = False
    tmrSlide.Enabled = True

End Sub

Private Sub Copy(Cntl As Control)

    Load Cntl
    With Cntl
        .Visible = True
        .Top = .Top + picViewport.Height
    End With 'CNTL

End Sub

Public Property Let Copyright(ForeColor As Long, nuCopyright As String)

    lbCopyright(0).Visible = True
    lbCopyright(0).ForeColor = ForeColor
    lbCopyright(0) = nuCopyright
    Copy lbCopyright(1)

End Property

Private Sub Form_Load()

  Dim i As Long

    picCurtain(0).Move 0, 0, picViewport.Width / 2, picViewport.Height
    picCurtain(1).Move picViewport.Width / 2, 0, picViewport.Width / 2, picViewport.Height
    picSlide.Move 0, 0, picViewport.Width, 2 * picViewport.Height
    With tmrSlide
        OriginalSpeed = .Interval
        .Enabled = True
    End With 'TMRSLIDE
    Theme = 13
    i = btHold.Height - 10
    hRgn1 = CreateRoundRectRgn(10, 10, i, i, i, i)
    SetWindowRgn btHold.hWnd, hRgn1, True
    hRgn2 = CreateRoundRectRgn(10, 10, i, i, i, i)
    SetWindowRgn btExit.hWnd, hRgn2, True
    Sidestep = -1

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        ReleaseCapture 'release the Mouse
        SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'non-client area button down (in caption)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    btExit.BackColor = NormalDark
    btHold.BackColor = NormalDark

End Sub

Private Sub Form_Paint()

  Dim Colr  As Long
  Dim Idx   As Long
  Dim w     As Long
  Dim h     As Long

    BackColor = TranslatedTheme
    With CT
        NormalDark = RGB(220 - .r, 220 - .g, 220 - .b)
        DarkerDark = RGB(200 - .r, 200 - .g, 200 - .b)
        btExit.BackColor = NormalDark
        btHold.BackColor = NormalDark
        lbDivider(0).BackColor = NormalDark
        w = picCurtain(0).ScaleWidth
        h = picCurtain(0).ScaleHeight
        For Idx = 0 To w
            Colr = 255 - Abs(128 - (Idx * 9) Mod 250)
            Colr = RGB(Colr - .r, Colr - .g, Colr - .b)
            picCurtain(0).Line (Idx, 0)-(Idx, h), Colr
            picCurtain(1).Line (w - Idx, 0)-(w - Idx, h), Colr
        Next Idx
        SetRect WindowRect, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels)
        For Idx = 0 To 255 Step 18
            Colr = 255 - Abs(128 - Idx)
            ForeColor = RGB(Colr - .r, Colr - .g, Colr - .b)
            With WindowRect
                Rectangle hDC, .Left, .Top, .Right, .Bottom
            End With 'WINDOWRECT
            InflateRect WindowRect, -1, -1
        Next Idx
    End With 'CT
    On Error Resume Next
        Copy lbDivider(1)
    On Error GoTo 0
    DoEvents
    Sleep 800

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If hRgn1 Then
        DeleteObject hRgn1
    End If
    If hRgn2 Then
        DeleteObject hRgn2
    End If

End Sub

Private Sub lbCopyright_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lbDivider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lbOtherstuff1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lbOtherstuff2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lbTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lbVersion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Function MakeColorTriple(r As Long, g As Long, b As Long) As ColorTriple

    With MakeColorTriple
        .r = r
        .g = g
        .b = b
    End With 'MAKECOLORTRIPLE

End Function

Public Property Let Otherstuff1(ForeColor As Long, nuOtherstuff1 As String)

    lbOtherstuff1(0).Visible = True
    lbOtherstuff1(0).ForeColor = ForeColor
    lbOtherstuff1(0) = nuOtherstuff1
    Copy lbOtherstuff1(1)

End Property

Public Property Let Otherstuff2(ForeColor As Long, nuOtherstuff2 As String)

    lbOtherstuff2(0).Visible = True
    lbOtherstuff2(0).ForeColor = ForeColor
    lbOtherstuff2(0) = nuOtherstuff2
    Copy lbOtherstuff2(1)

End Property

Private Sub picIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub picSlide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Public Property Let Theme(nuColor As Long)

    Select Case nuColor
      Case Is < 1, Is > 27
      Case Else
        myThemeColor = nuColor
    End Select

End Property

Public Property Let Title(ForeColor As Long, nuTitle As String)

    lbTitle(0).Visible = True
    lbTitle(0).ForeColor = ForeColor
    lbTitle(0) = nuTitle
    Copy lbTitle(1)

End Property

Private Sub tmrSlide_Timer()

    With picSlide
        If .Top = -picViewport.Height Then
            .Top = 0
          Else 'NOT .TOP...
            .Top = .Top - ScaleY(1, vbPixels, ScaleMode)
        End If
    End With 'PICSLIDE
    With picCurtain(0)
        .Left = .Left + Sidestep
    End With 'PICCURTAIN(0)
    With picCurtain(1)
        .Left = .Left - Sidestep
        If .Left >= picViewport.Width Or .Left <= picViewport.Width / 2 Then
            Sidestep = 0
        End If
    End With 'PICCURTAIN(1)
    With tmrSlide
        If MouseIsDown Then
            If .Interval > 100 Then
                .Enabled = False
              Else 'NOT .INTERVAL...
                .Interval = .Interval * 1.1
            End If
          Else 'MOUSEISDOWN = FALSE/0
            If .Interval > OriginalSpeed Then
                .Interval = .Interval / 1.1
            End If
        End If
    End With 'TMRSLIDE

End Sub

Public Property Get TranslatedTheme() As Long

    CT = MakeColorTriple((myThemeColor \ 9) Mod 3, (myThemeColor \ 3) Mod 3, myThemeColor Mod 3)
    With CT
        .r = .r * 32
        .g = .g * 32
        .b = .b * 32
        TranslatedTheme = RGB(255 - .r, 255 - .g, 255 - .b)
    End With 'CT

End Property

Public Property Let Version(ForeColor As Long, nuVersion As String)

    lbVersion(0).Visible = True
    lbVersion(0).ForeColor = ForeColor
    lbVersion(0) = nuVersion
    Copy lbVersion(1)

End Property

':) Ulli's VB Code Formatter V2.21.6 (2006-Mrz-26 19:06)  Decl: 67  Code: 307  Total: 374 Lines
':) CommentOnly: 23 (6,1%)  Commented: 15 (4%)  Empty: 95 (25,4%)  Max Logic Depth: 4
