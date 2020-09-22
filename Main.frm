VERSION 5.00
Begin VB.Form Main_Form 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13590
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFF00&
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   906
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   2760
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   10800
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Border"
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   12
      Text            =   "0"
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   12000
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto"
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9360
      Top             =   2880
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dat"
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option3"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direção"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.CommandButton Turn 
      Caption         =   "Girar"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1860
      Index           =   0
      Left            =   10800
      Picture         =   "Main.frx":164A
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Roulette Menu it is a nice 3D Program for your Desktop.
'Allows choise the image and action for each Icone.
'4 Roulette types, with Zoom and Stretch features.

'MOUSE RIGHT BUTTOM to Show Popup Menu
'MOUSE LEFT BUTTON and DRAG to MOVE
'MOUSE WHEEL or UP/DOWN Keys to TURN
'+ or - KEY with SHIFT to Vertical Stretch
'+ or - KEY with SHIFT+CTRL to ZOOM
'DOUBLE CLICK on the ICONE to Run One Program, Send an E-Mail or Open an URL

'You can have several groups in same time.
'To make this:
'- Make the Executable file
'- Copy the EXE file and the BAG file
'- Paste them as new files
'- Rename them preserving the extensions
'     Eg: MyNewGroup1.EXE and MyNewGroup1.BAG
'- Run the EXE files individually

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
        ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE As Long = (-16&)
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOMOVE As Long = 2
Private Const SWP_NOSIZE As Long = 1

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type FAT
    X As Double
    Y As Double
End Type

Private Type POINT
    X As Integer
    Y As Integer
End Type

Private Fator As FAT

Private PT As POINTAPI

Public Mouse_Under As Integer
Private Amount_Groups As Integer
Private Group_Name() As String
Private Last_mouse As Integer
Private Color(0 To 11) As Long
Private Working As Boolean
Private hRgn As Long
Private Rel_W As Long
Private Rel_H As Long
Private Direction As Integer
Private Frame As Integer
Private Q(0 To 119, 1 To 4) As POINT
Private C(0 To 119) As Integer
Private Ordem(0 To 119) As Integer
Private XX As Long
Private YY As Long
Private Capture As Boolean
Private Declare Sub InitCommonControls Lib "comctl32" () ':) Line inserted by Formatter

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Check1_Click()

    Timer1.Enabled = Check1.Value = 1
    Combo1.SetFocus

End Sub

Private Sub Combo1_Click()

    If Combo1.ListIndex = 1 Then
        Exit Sub
    End If

    Select Case Combo1.ListIndex
      Case 0
        Option1(0).Value = True
      Case 2
        Option1(1).Value = True
    End Select

    Turn_Click
    Combo1.ListIndex = 1

End Sub

Private Sub Command1_Click()

    Visible = False

    SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Xor _
                                                    (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)

    SetWindowPos Me.hWnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or _
                 SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED

    Visible = True
    Combo1.SetFocus

End Sub

Private Sub Form_Activate()

    Combo1.SetFocus

End Sub

Private Sub Form_DblClick()

    If Icones(Last_Choised).Shell <> "" Then
        ShellExecute hWnd, "open", Icones(Last_Choised).Shell, vbNullString, vbNullString, 1
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim X As Long
  Dim Y As Long

    X = Width
    Y = Height

    Select Case KeyCode
      Case 107
        X = X - (X / 10) * ((Shift = 2) Or (Shift = 3))
        Y = Y - (Y / 10) * ((Shift = 1) Or (Shift = 3))
        If X > 1000 And Y > 1000 And X < Screen.Width And Y < Screen.Height Then
            Move Left, Top, X, Y
        End If

      Case 109
        X = X + (X / 10) * ((Shift = 2) Or (Shift = 3))
        Y = Y + (X / 10) * ((Shift = 1) Or (Shift = 3))
        If X > 1000 And Y > 1000 And X < Screen.Width And Y < Screen.Height Then
            Move Left, Top, X, Y
        End If
      Case 27
        If Shift Then
            Unload Menu_Form
            Unload Mask_Form
            Unload Properties_Form
            Unload Me
            Exit Sub
        End If

    End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu Menu_Form.Menu
        If End_Prg Then
            Unload Me
            Exit Sub
        End If
        Draw
        Exit Sub
    End If

    GetCursorPos PT
    XX = PT.X * Screen.TwipsPerPixelX - Left
    YY = PT.Y * Screen.TwipsPerPixelY - Top
    Capture = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim i As Integer

    If Capture And Button = 1 Then
        DoEvents
        GetCursorPos PT
        Move (PT.X * Screen.TwipsPerPixelX - XX), (PT.Y * Screen.TwipsPerPixelY - YY)
    End If
    For i = 0 To 11
        If Mask_Form.POINT(X, Y) = Color(i) Then
            Mouse_Under = i
            Last_Choised = i
            If Last_mouse <> Mouse_Under Then
                Draw
                If Menu_Form.Auto_Turn.Checked Then
                    Timer1.Enabled = False
                    Timer2.Enabled = True
                End If
            End If
            Last_mouse = Mouse_Under
            Timer3.Enabled = True
            Exit Sub
        End If
    Next i
    Mouse_Under = -1

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Capture = False

End Sub

Private Sub Form_Resize()

    Do While Working
        DoEvents
    Loop

    Fator.X = Width / Rel_W
    Fator.Y = Height / Rel_H

    Draw

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Save_Bag

End Sub

Private Sub Timer1_Timer()

    Turn_Click

End Sub

Private Sub Timer2_Timer()

    Check1.Value = 1
    Timer1.Enabled = True
    Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()

  Dim PX As Long
  Dim PY As Long

    GetCursorPos PT
    PX = PT.X - (Left / Screen.TwipsPerPixelX)
    PY = PT.Y - (Top / Screen.TwipsPerPixelY)

    If Mask_Form.POINT(PX, PY) = 14215660 Or Mask_Form.POINT(PX, PY) = -1 Then
        Timer3.Enabled = False
        Mouse_Under = -1
        Last_mouse = -2
        Draw
    End If

End Sub

Private Sub Turn_Click()

  Static X As Integer

  Dim i As Integer
  Dim k As Integer

    For i = 0 To 11
        C(i) = C(i) + Direction

        If C(i) > 118 Then
            C(i) = 0
        End If

        If C(i) < 0 Then
            C(i) = 118
        End If
    Next i

    Draw

End Sub

Private Sub Form_Load()

  Dim X As Integer
  Dim i As Integer
  Dim Ret As Long
  Dim objBag            As New PropertyBag
  Dim vntBagContents    As Variant

    For i = 1 To 11
        Load Pic(i)
    Next i

    Open App.Path & "\" & App.EXEName & ".Bag" For Binary As 1
    Get #1, , vntBagContents
    Close 1

    With objBag
        .Contents = vntBagContents
        For i = 0 To 11
            Pic(i).Picture = .ReadProperty("Picture" & CStr(i))
            Icones(i).Shell = .ReadProperty("Action" & CStr(i))
            Icones(i).Transp_Color = .ReadProperty("Transp_Color" & CStr(i))
            Icones(i).Type = .ReadProperty("Type" & CStr(i))
        Next i
    End With

    For i = 0 To 11
        Color(i) = QBColor(i)
    Next i

    BackColor = &HFF
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes hWnd, &HFF, 255, LWA_COLORKEY Or LWA_ALPHA

    Combo1.AddItem 1
    Combo1.AddItem 0
    Combo1.AddItem -1
    Combo1.Move -1000
    C(0) = 0
    C(1) = 10
    C(2) = 20
    C(3) = 30
    C(4) = 40
    C(5) = 50
    C(6) = 60
    C(7) = 70
    C(8) = 80
    C(9) = 90
    C(10) = 100
    C(11) = 110

    Direction = 1

    For i = 0 To 59
        Ordem(X) = i
        Ordem(X + 1) = 118 - i
        X = X + 2
    Next i

    Option2_Click (objBag.ReadProperty("Option"))
    Move objBag.ReadProperty("Left"), objBag.ReadProperty("Top"), objBag.ReadProperty("Width"), objBag.ReadProperty("height")

End Sub

Private Sub Option1_Click(index As Integer)

    If Option1(0).Value Then
        Direction = 1
      Else
        Direction = -1
    End If

End Sub

Private Sub Option2_Click(index As Integer)

  Dim B() As Byte

    Choised_type = index
    Select Case index
      Case 0
        Rel_W = 4000
        Rel_H = 9800
        Width = 4000
        Height = 9800

      Case 1
        Rel_W = 4500
        Rel_H = 9800
        Width = 4500
        Height = 9800

      Case 2
        Rel_W = 9500
        Rel_H = 3000
        Width = 9500
        Height = 3000

      Case 3
        Rel_W = 9500
        Rel_H = 3000
        Width = 9500
        Height = 3000

    End Select

    B = LoadResData(index + 1, "Rotation")

    CopyMemory ByVal VarPtr(Q(0, 1)), ByVal VarPtr(B(0)), UBound(B) - 1

    Turn_Click

    If Combo1.Visible Then
        Combo1.SetFocus
    End If

End Sub

Public Sub DrawPic(index As Integer, icone)

  Dim StartX As Long
  Dim StartY As Long
  Dim StartHeight As Long
  Dim EndHeight As Long
  Dim OutWidth As Long
  Dim OutYOffset As Long

    Working = True

    StartX = Q(index, 1).X * Fator.X
    StartY = Q(index, 1).Y * Fator.Y
    StartHeight = (Q(index, 3).Y - Q(index, 1).Y) * Fator.Y
    EndHeight = (Q(index, 4).Y - Q(index, 2).Y) * Fator.Y
    OutWidth = (Q(index, 2).X - Q(index, 1).X) * Fator.X
    OutYOffset = (Q(index, 2).Y - Q(index, 1).Y) * Fator.X
    Picture2.Width = OutWidth
    Picture2.Height = StartHeight
    Picture1.Width = OutWidth
    Picture1.Height = StartHeight
    Picture2.Cls
    PerspBltX Picture2.hdc, 0, 0, OutWidth, StartHeight, EndHeight, OutYOffset, Pic(icone).hdc, Pic(icone).ScaleWidth, Pic(icone).ScaleHeight

    If Mouse_Under = icone Then
        BitBlt Picture2.hdc, 0, 0, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, vbSrcAnd
    End If

    GdiTransparentBlt hdc, StartX, StartY, Picture2.Width, Picture2.Height, Picture2.hdc, 0, 0, OutWidth, StartHeight, Icones(icone).Transp_Color

    Mask_Form.Line (Q(index, 1).X * Fator.X, Q(index, 1).Y * Fator.Y)-(Q(index, 4).X * Fator.X, Q(index, 4).Y * Fator.Y), QBColor(icone), BF

    Working = False

End Sub

Private Sub Draw()

  Dim i As Integer
  Dim k As Integer

    Cls

    Mask_Form.Cls
    Mask_Form.Width = Width
    Mask_Form.Height = Height

    'On Error Resume Next
    For i = 0 To 118
        For k = 0 To 11
            If Ordem(i) = C(k) Then
                Do While Working
                    DoEvents
                Loop
                DrawPic C(k), k
                Exit For
            End If
        Next k
    Next i

    Refresh

End Sub

Private Sub Save_Bag()

  Dim objBag            As New PropertyBag
  Dim vntBagContents    As Variant
  Dim i As Integer

    With objBag
        For i = 0 To 11
            .WriteProperty "Picture" & CStr(i), Pic(i)
            .WriteProperty "Type" & CStr(i), Icones(i).Type
            .WriteProperty "Action" & CStr(i), Icones(i).Shell
            .WriteProperty "Transp_Color" & CStr(i), Icones(i).Transp_Color
        Next i
        .WriteProperty "Left", Left
        .WriteProperty "Top", Top
        .WriteProperty "Width", Width
        .WriteProperty "Height", Height
        .WriteProperty "Option", Choised_type
        .WriteProperty "Groups", Amount_Groups
        For i = 1 To Amount_Groups
            .WriteProperty "Group" & CStr(i), Group_Name(i)
        Next i
        vntBagContents = .Contents
    End With

    Open App.Path & "\" & App.EXEName & ".Bag" For Binary As 1
    Put #1, , vntBagContents
    Close 1

End Sub

