VERSION 5.00
Begin VB.Form Properties_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "Properties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   6360
      Picture         =   "Properties.frx":164A
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image"
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1860
         Left            =   120
         Picture         =   "Properties.frx":CCAC
         ScaleHeight     =   124
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   9
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Send an E-mail"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Open an URL"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Open a Program"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Transparent Color"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "Properties_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Selected_item As Integer
Private Selected_Image As String

Private Sub Command1_Click()

    Text1.Text = FileDialog(Me, False, "Open Program", "Exe|*.exe")

End Sub

Private Sub Command2_Click()

    If Selected_item = 2 Then
        Icones(Item).Shell = "mailto:" + Text1.Text
      Else
        Icones(Item).Shell = Text1.Text
    End If
    Icones(Item).Type = Selected_item
    Icones(Item).Transp_Color = Label1.BackColor
    Main_Form.Pic(Item).Picture = Picture1.Picture
    Unload Me

End Sub

Private Sub Command3_Click()

    Unload Me

End Sub

Private Sub Command4_Click()

  Dim X As String

    X = FileDialog(Me, False, "Set Picture", "Bmp;Gif;Jpg|*.bmp;*Gif;*.jpg")

    If X <> "" Then
        Picture2.Picture = LoadPicture(X)
        SetStretchBltMode Picture1.hdc, COLORONCOLOR
        StretchBlt Picture1.hdc, 0, 0, Picture1.Width, Picture1.Height, Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, vbSrcCopy
        Picture1.Picture = Picture1.Image
        Picture1.Refresh
        Picture2.Picture = LoadPicture()
    End If

End Sub

Private Sub Form_Load()

    Item = Last_Choised
    Option1(Icones(Item).Type).Value = True
    Text1 = Icones(Item).Shell
    Selected_item = Icones(Item).Type
    Picture1.Picture = Main_Form.Pic(Item)
    Label1.BackColor = Icones(Item).Transp_Color
    Label3.Caption = Right("00000000" & Hex(Icones(Item).Transp_Color), 8)

End Sub

Private Sub Option1_Click(index As Integer)

    Selected_item = index
    Text1.Top = Option1(index).Top

    Select Case index
      Case 0
        Command1.Visible = True
        Text1.Text = ""

      Case 1
        Command1.Visible = False
        Text1.Text = ""

      Case 2
        Command1.Visible = False
        Text1.Text = ""

    End Select

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label1.BackColor = Picture1.POINT(X, Y)
    Label3.Caption = Right("00000000" & Hex(Picture1.POINT(X, Y)), 8)

End Sub

