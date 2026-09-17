VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopyMove 
   Caption         =   "Choose Action"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3990
   OleObjectBlob   =   "frmCopyMove.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCopyMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedAction As String

Private Sub cmdCopy_Click()
    SelectedAction = "COPY"
    Me.Hide
End Sub

Private Sub cmdMove_Click()
    SelectedAction = "MOVE"
    Me.Hide
End Sub

Private Sub UserForm_Click()

End Sub
