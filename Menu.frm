VERSION 5.00
Begin VB.Form Menu_Form 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Properties 
         Caption         =   "Proprieties"
      End
      Begin VB.Menu Type_show 
         Caption         =   "Type"
         Begin VB.Menu Type_index 
            Caption         =   "Vertical Left"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu Type_index 
            Caption         =   "Vertical Right"
            Index           =   1
         End
         Begin VB.Menu Type_index 
            Caption         =   "Horizontal_Up"
            Index           =   2
         End
         Begin VB.Menu Type_index 
            Caption         =   "Horizontal Down"
            Index           =   3
         End
      End
      Begin VB.Menu Auto_Turn 
         Caption         =   "Turn On"
      End
      Begin VB.Menu Border 
         Caption         =   "Border"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Begin VB.Menu About_index 
            Caption         =   "Autor    :     Agustin Rodriguez"
            Index           =   0
         End
         Begin VB.Menu About_index 
            Caption         =   "Year      :     2011"
            Index           =   1
         End
         Begin VB.Menu About_index 
            Caption         =   "Country:     Brazil"
            Index           =   2
         End
         Begin VB.Menu About_index 
            Caption         =   "E-Mail    :    virtual_guitar_1@hotmail.com"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Auto_Turn_Click()

    Auto_Turn.Checked = Auto_Turn.Checked Xor -1
    Main_Form.Check1.Value = Abs(Auto_Turn.Checked)
    Main_Form.Timer2.Enabled = False

End Sub

Private Sub Exit_Click()

    Unload Mask_Form
    Unload Properties_Form
    End_Prg = True
    Unload Me

End Sub

Private Sub Border_Click()

    Border.Checked = Border.Checked Xor -1
    Main_Form.Command1 = True

End Sub

Private Sub Properties_Click()

    Properties_Form.Show 1

End Sub

Private Sub Type_index_Click(index As Integer)

  Static Last_index As Integer

    Type_index(Last_index).Checked = False
    Type_index(index).Checked = True
    Main_Form.Option2(index).Value = True
    Last_index = index

End Sub

