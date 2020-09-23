VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "update"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "0123456789"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox segs 
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Text            =   "10"
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   450
      Left            =   0
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "text"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label se 
      Caption         =   "segs"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Cls
For asdf = 0 To CInt(segs.Text) - 1
DrawSegment 20 * asdf, 0, OFFColor, ONColor, Picture1
Next
For x = 0 To Len(Text1.Text) - 1
Number2Segment 21 * x, 0, Mid(Text1.Text, x + 1, 1), Picture1
Next
End Sub
