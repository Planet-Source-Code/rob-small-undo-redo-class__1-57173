VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Undo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' **********************************************************************************************
' Class     : Undo
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : An instance of this class holds a single undo command
' **********************************************************************************************

Public Field As Control
Public PrevVal As Variant
Public NewVal As Variant

