VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mark 
   Caption         =   "附图标记"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   OleObjectBlob   =   "mark.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "mark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_Click()
    Call step3
    Call step4
    Unload Me
End Sub
Private Sub type1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        a = type1.Text
        If Len(a) > 0 Then
            With type1
                .Text = ""
                .Visible = False
            End With
            Call step2
            With type2
                .Visible = True
                .SetFocus
            End With
        Else
            Call step3
            Call step4
            Unload Me
        End If
    End If
End Sub
Private Sub type2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        If KeyCode = vbKeyReturn Then
        a = type2.Text
        If Len(a) > 0 Then
            With type2
                .Text = ""
                .Visible = False
            End With
            Call step2
            With type1
                .Visible = True
                .SetFocus
            End With
        Else
            Call step3
            Call step4
            Unload Me
        End If
    End If
End Sub
Private Sub UserForm_Initialize()
    With Me
'        .StartUpPosition = 0
'        .Left = 650
'        .Top = 200
        .type1.Visible = True
        .type2.Visible = False
    End With
End Sub
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    Call step3
    Call step4
End Sub
