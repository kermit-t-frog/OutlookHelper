VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PWBox 
   Caption         =   "Password"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4050
   OleObjectBlob   =   "PWBox.frx":0000
   StartUpPosition =   1  
End
Attribute VB_Name = "PWBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private value_ As String

Private Sub btn_Cancel_Click()
    Me.hide
End Sub

Public Property Get Value() As String
    Value = value_
End Property

Private Sub btn_OK_Click()
    value_ = Me.txt_Input.Value
    Me.hide
End Sub
