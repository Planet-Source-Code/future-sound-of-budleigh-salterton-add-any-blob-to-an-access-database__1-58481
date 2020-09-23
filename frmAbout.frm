VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVersion 
      Caption         =   "Label1"
      Height          =   735
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3255
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub



Private Sub picMe_Click()
    Unload Me
End Sub
