VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Barcode Reader Test Form"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "i2 of 5"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Codabar"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "3 of 9"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "2"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "5"
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "Result"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Barcode Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Verbos"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Retry"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''
'' BARCODE READER - Test Form              ''
'' By Paul Bahlawan Aug 2003               ''
''   -updated nov 2011                     ''
'''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim bcType As Long

Private Sub Form_Load()
    Picture1.Picture = LoadPicture(App.Path & "\Test files\3 bc types.bmp")
    Picture1.ScaleMode = vbPixels
End Sub

'select barcode type
Private Sub Option1_Click(Index As Integer)
    bcType = Index
End Sub

'read barcode
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Text3.Text = bcRead(Picture1, CLng(x), CLng(y), bcType, Val(Text1.Text), Val(Text2.Text))
End Sub

'drag and drop image
Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If LCase(Right$(Data.Files(1), 4)) <> ".bmp" Then Exit Sub
    
    Picture1.Picture = LoadPicture(Data.Files(1))
    Text3.Text = ""
End Sub

'resize form to picturebox
Private Sub Picture1_Resize()
    Form1.Width = Picture1.Width + 350
    Form1.Height = Picture1.Height + Picture1.Top + 600

End Sub
