VERSION 5.00
Begin VB.Form FrmPng 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   120
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FrmPng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lWindows As LayeredWindow

Public Sub LoadPng(x As Integer)
Set lWindows = New LayeredWindow
'lWindows.MakeTrans App.Path & "\ball\" & x & ".png", Me
lWindows.MakeTrans App.Path & "\love\" & x & ".png", Me
Set lWindows = Nothing
End Sub
