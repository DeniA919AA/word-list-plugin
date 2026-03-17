VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents CloseButton As MSForms.CommandButton
Attribute CloseButton.VB_VarHelpID = -1
Private ResultBox As MSForms.TextBox

Private Sub UserForm_Initialize()
    Me.Caption = "List Detection Results"
    Me.Width = 700
    Me.Height = 550

    Set ResultBox = Me.Controls.Add("Forms.TextBox.1", "ResultBox")
    ResultBox.Multiline = True
    ResultBox.ScrollBars = 2
    ResultBox.WordWrap = False
    ResultBox.Left = 10
    ResultBox.Top = 10
    ResultBox.Width = 660
    ResultBox.Height = 460
    ResultBox.Font.Name = "Courier New"
    ResultBox.Font.Size = 10
    ResultBox.Text = ResultText

    Set CloseButton = Me.Controls.Add("Forms.CommandButton.1", "CloseBtn")
    CloseButton.Caption = "Close"
    CloseButton.Left = 300
    CloseButton.Top = 480
    CloseButton.Width = 80
    CloseButton.Height = 25
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub
