VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unicode TextEdit sample"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2895
      Left            =   4560
      Max             =   5
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1800
      List            =   "Form1.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Multiline"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin Project1.TextEdit edit2 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
      _extentx        =   7646
      _extenty        =   3413
      multiline_      =   -1  'True
      text_           =   $"Form1.frx":0047
   End
   Begin Project1.TextEdit edit1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _extentx        =   7646
      _extenty        =   661
      multiline_      =   0   'False
      text_           =   "Ïðèâåò!"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    Select Case Combo1.ListIndex
    Case 0:
        edit2.BackColor = vbRed
        edit2.ForeColor = vbGreen
        edit1.BackColor = vbGreen
        edit1.ForeColor = vbRed
    Case 1:
        edit2.BackColor = vbBlue
        edit2.ForeColor = vbCyan
        edit1.BackColor = vbCyan
        edit1.ForeColor = vbBlue
    Case 2:
        edit2.BackColor = vbMagenta
        edit2.ForeColor = vbYellow
        edit1.BackColor = vbYellow
        edit1.ForeColor = vbMagenta
    Case 3:
        edit2.BackColor = vbWhite
        edit2.ForeColor = vbBlack
        edit1.BackColor = vbBlack
        edit1.ForeColor = vbWhite
    End Select
End Sub

Private Sub Command1_Click()
    edit2.Multiline = Not edit2.Multiline
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 1
    VScroll1.Value = 1
    edit1.text = ChrW$(161) & "Hola, mundo! http://sf.net/projects/audica"
End Sub

Private Sub VScroll1_Change()
    edit1.Frame(vbButtonShadow) = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    Call VScroll1_Change
End Sub
