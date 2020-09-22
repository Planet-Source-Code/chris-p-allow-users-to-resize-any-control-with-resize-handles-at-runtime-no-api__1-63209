Attribute VB_Name = "mod_resizer"
'RESIZE CONTROLS DURING RUNTIME
'By: CHRIS_P
'This simple module will allow you to resize any control during runtime
'1 - make a 6x6 pixel picture box on a form and call it "handle" and give it an index value of 0 (so that it becomes a control array)
'2 - add the following code to the form
    'Private Sub handle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '    handle_press X, Y
    'End Sub
    '
    'Private Sub handle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '    handle_move Index, Button, Shift, X, Y, Me
    'End Sub
'3 - allow users to resize or move any control on the form you want using the following function:
'   allowresize <control name>,<form where control is located>

Public handlex As Single
Public handley As Single
Public SelectedControl As Control


Sub handles_init(HandleContainer As Object, ParentForm As Form)
    
    Set ParentForm.handle(0).Container = HandleContainer
    ParentForm.handle(0).Visible = False
    ParentForm.handle(0).MousePointer = 8
    For i = 1 To 8
        If ParentForm.handle.Count = i Then Load ParentForm.handle(i)
        
        Select Case (i)
            Case 1
                ParentForm.handle(1).MousePointer = 7 'N-S
            Case 2
                ParentForm.handle(2).MousePointer = 6 'NE-SW
            Case 3
                ParentForm.handle(3).MousePointer = 9 ' E-W
            Case 4
                ParentForm.handle(4).MousePointer = 8 ' NW-SE
            Case 5
                ParentForm.handle(5).MousePointer = 7 'NS
            Case 6
                ParentForm.handle(6).MousePointer = 6 'NE-SW
            Case 7
                ParentForm.handle(7).MousePointer = 9
            Case 8
                ParentForm.handle(8).MousePointer = 15
        End Select
        Set ParentForm.handle(i).Container = HandleContainer
            ParentForm.handle(i).Visible = False
    Next
    
End Sub
Sub handles_hide(ParentForm As Form)
On Error Resume Next
For i = 0 To 8
        If (ParentForm.handle(i).Visible = True) Then ParentForm.handle(i).Visible = False
Next
End Sub
Sub handles_show(ctl As Control, frm As Form)

For i = 0 To 8
        If (frm.handle(i).Visible = False) Then frm.handle(i).Visible = True
Next
frm.handle(0).Move ctl.Left - frm.handle(0).Width, ctl.Top - frm.handle(0).Height
frm.handle(1).Move ctl.Left + ((ctl.Width / 2) - frm.handle(1).Width / 2), ctl.Top - frm.handle(0).Height
frm.handle(2).Move ctl.Left + ctl.Width, ctl.Top - frm.handle(0).Height
frm.handle(3).Move ctl.Left + ctl.Width, ctl.Top + ((ctl.Height / 2) - frm.handle(3).Height / 2)
frm.handle(4).Move ctl.Left + ctl.Width, ctl.Top + ctl.Height
frm.handle(5).Move ctl.Left + ((ctl.Width / 2) - frm.handle(1).Width / 2), ctl.Top + ctl.Height
frm.handle(6).Move ctl.Left - frm.handle(6).Width, ctl.Top + ctl.Height
frm.handle(7).Move ctl.Left - frm.handle(7).Width, ctl.Top + ((ctl.Height / 2) - frm.handle(7).Height / 2)
frm.handle(8).Move ctl.Left + ((ctl.Width / 2) - frm.handle(8).Width / 2), ctl.Top + ((ctl.Height / 2) - frm.handle(8).Height / 2)

End Sub

Sub handle_move(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, frm As Form)
Static prevx, prevy
On Error Resume Next
If Button = 1 Then

        If Abs(X - prevx) < 1 And Abs(Y - prevy) < 1 Then Exit Sub
        prevy = Y: prevx = X
        
    Select Case Index
        Case 0
            frm.handle(Index).Move frm.handle(Index).Left + (X - handlex), frm.handle(Index).Top + (Y - handley)
            SelectedControl.Move SelectedControl.Left + (X - handlex), SelectedControl.Top + (Y - handley), SelectedControl.Width - (X - handlex), SelectedControl.Height - (Y - handley)
            handles_show SelectedControl, frm
        Case 1
            frm.handle(Index).Move frm.handle(Index).Left, frm.handle(Index).Top + (Y - handley)
            SelectedControl.Move SelectedControl.Left, SelectedControl.Top + (Y - handley), SelectedControl.Width, SelectedControl.Height - (Y - handley)
            handles_show SelectedControl, frm
        Case 2
            frm.handle(Index).Move frm.handle(Index).Left + (X - handlex), frm.handle(Index).Top + (Y - handley)
            SelectedControl.Move SelectedControl.Left, SelectedControl.Top + (Y - handley), SelectedControl.Width + (X - handlex), SelectedControl.Height - (Y - handley)
            handles_show SelectedControl, frm
        Case 3
            frm.handle(Index).Left = frm.handle(Index).Left + (X - handlex)
            SelectedControl.Width = SelectedControl.Width + (X - handlex)
            handles_show SelectedControl, frm
        Case 4
            frm.handle(Index).Move frm.handle(Index).Left + (X - handlex), frm.handle(Index).Top - (Y - handley)
            SelectedControl.Move SelectedControl.Left, SelectedControl.Top, SelectedControl.Width + (X - handlex), SelectedControl.Height + (Y - handley)
            handles_show SelectedControl, frm
        Case 5
            frm.handle(Index).Move frm.handle(Index).Left, frm.handle(Index).Top - (Y - handley)
            SelectedControl.Height = SelectedControl.Height + (Y - handley)
            handles_show SelectedControl, frm
        Case 6
            frm.handle(Index).Move frm.handle(Index).Left - (X - handlex), frm.handle(Index).Top - (Y - handley)
            SelectedControl.Move SelectedControl.Left + (X + handlex), SelectedControl.Top, SelectedControl.Width - (X + handlex), SelectedControl.Height + (Y - handley)
            handles_show SelectedControl, frm
        Case 7
            frm.handle(Index).Left = frm.handle(Index).Left - (X - handlex)
            SelectedControl.Left = SelectedControl.Left + (X + handlex)
            SelectedControl.Width = SelectedControl.Width - (X + handlex)
            handles_show SelectedControl, frm
        Case 8
            frm.handle(Index).Move frm.handle(Index).Left - (X - handlex), frm.handle(Index).Top - (Y - handley)
            SelectedControl.Move SelectedControl.Left + (X - handlex), SelectedControl.Top + (Y - handley)
            handles_show SelectedControl, frm
        
    End Select
End If

End Sub

Sub allowresize(ctl As Control, frm As Form)
    handles_init ctl.Container, frm
    Set SelectedControl = ctl
    handles_show ctl, frm

End Sub

Sub handle_press(X As Single, Y As Single)
    handlex = X
    handley = Y
End Sub
