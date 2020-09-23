VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6930
   ClientLeft      =   5820
   ClientTop       =   5250
   ClientWidth     =   9735
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
  Form1.Show
End Sub

Private Sub mnuRedo_Click()
  ' Call the Redo function on Form1
  Call Form1.LocalUndo.RedoEdit
End Sub

Private Sub mnuUndo_Click()
  ' Call the Undo function on Form1
  Call Form1.LocalUndo.UndoEdit
End Sub
