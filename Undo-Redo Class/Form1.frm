VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   7260
   ClientTop       =   6570
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   5400
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1440
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LocalUndo As New UndoWrapper

Private Sub Form_Load()
  ' Set up the class to allow Undo on the current form.  mnuUndo and mnuRedo
  ' are the Undo and Redo menu items on the MDI form.  These parameters are
  ' not essential since your application may not use a menu.
  Call LocalUndo.SetupClass(Me, MDIForm1.mnuUndo, MDIForm1.mnuRedo)

End Sub

Private Sub Text1_Change()
  ' If the contents of the Textbox change, clear the history
  ' from this point forward
  Call LocalUndo.ClearHistoryTail
End Sub

Private Sub Text1_GotFocus()
  ' When the Textbox gets focus, Add its details to the History list
  Call LocalUndo.AddUndoHistory("Text1")
End Sub

Private Sub Text2_Change()
  Call LocalUndo.ClearHistoryTail
End Sub

Private Sub Text2_GotFocus()
  Call LocalUndo.AddUndoHistory("Text2")
End Sub

Private Sub Combo1_Click()
  Call LocalUndo.ClearHistoryTail
End Sub

Private Sub Combo1_GotFocus()
  Call LocalUndo.AddUndoHistory("Combo1")
End Sub

Private Sub Check1_Click()
  Call LocalUndo.ClearHistoryTail
End Sub

Private Sub Check1_GotFocus()
  Call LocalUndo.AddUndoHistory("Check1")
End Sub

Private Sub Check2_Click()
  Call LocalUndo.ClearHistoryTail
End Sub

Private Sub Check2_GotFocus()
  Call LocalUndo.AddUndoHistory("Check2")
End Sub



