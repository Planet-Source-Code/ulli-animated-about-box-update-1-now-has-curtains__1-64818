VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Test"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3630
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btAbout 
      Caption         =   "About Box"
      Height          =   645
      Left            =   1163
      TabIndex        =   0
      Top             =   645
      Width           =   1305
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btAbout_Click()

    Load frmAbout '!!!

    With frmAbout

        .Theme = Timer Mod 27 + 1 'for testing - choose a fixed theme (1..27) that suits you

        'after the theme is set the translated theme color is available in .TranslatedTheme
        .AppIcon(.TranslatedTheme) = Icon

        .Copyright(vbYellow) = "Enter your copyright here"

        .Otherstuff1(vbRed) = "Enter other stuff here"

        .Otherstuff2(vbYellow) = "You may enter a longer desription of the project or anything else which spans several lines here"

        .Title(.TranslatedTheme) = "Testing the About Box"

        .Version(&HC0C0A0) = "This is for the version"

        .Show vbModal, Me
        '___________________________

        'that's all

        '¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
    End With 'FRMABOUT

End Sub

Private Sub Form_Load()

    Randomize Timer 'to create a random theme color

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Mrz-26 17:06)  Decl: 1  Code: 40  Total: 41 Lines
':) CommentOnly: 4 (9,8%)  Commented: 4 (9,8%)  Empty: 17 (41,5%)  Max Logic Depth: 2
