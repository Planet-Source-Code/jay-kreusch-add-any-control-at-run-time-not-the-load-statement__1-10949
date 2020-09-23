<div align="center">

## Add ANY control at run\-time \(NOT the Load Statement\!\)


</div>

### Description

Adds a control to your form at run-time. Does not use control arrays or the Load statement. Control does not even need to be referenced. NO API.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jay Kreusch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jay-kreusch.md)
**Level**          |Beginner
**User Rating**    |4.8 (155 globes from 32 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jay-kreusch-add-any-control-at-run-time-not-the-load-statement__1-10949/archive/master.zip)





### Source Code

```
Option Explicit
' If you are adding an ActiveX control at run-time that is
' not referenced in your project, you need to declare it
' as VBControlExtender.
Dim WithEvents ctlDynamic As VBControlExtender
Dim WithEvents ctlCommand As VB.CommandButton
Dim WithEvents ctlText As VB.TextBox
Private Sub ctlCommand_Click()
  'since we delcared withevents, we can use them
  ctlDynamic.object.Value = CDate(ctlText.Text)
End Sub
Private Sub ctlDynamic_ObjectEvent(Info As EventInfo)
  'This is sort of an 'all-in-one' event
  'so you have to check parameters and event name
  Dim p As EventParameter
  Debug.Print Info.Name
  For Each p In Info.EventParameters
    Debug.Print p.Name, p.Value
  Next
  Select Case Info.Name
    Case "NewMonth"
      ctlText.Text = ctlDynamic.object.Value
    Case "Click"
      MsgBox ctlDynamic.object.Value
  End Select
End Sub
Private Sub Form_Load()
  'If you get run-time error number 732.
  'Then the control isn't in the liscenses collection
  'Use this line with the ProgID you want
  'Licenses.Add [ProgID]
  ' Add a control and set the properties of the control
  Set ctlDynamic = Controls.Add("mscal.calendar", "calMain", Form1)
  With ctlDynamic
    .Move 1, 400, 4000, 3000
    .Visible = True
  End With
  ' add a textbox and set properties for the textbox
  Set ctlText = Controls.Add("VB.TextBox", "ctlText1", Form1)
  With ctlText
    .Move 1, 1, 3400, 100
    .Text = ctlDynamic.object.Value
    .Visible = True
  End With
  ' Add a CommandButton.
  Set ctlCommand = Controls.Add("VB.CommandButton", "ctlCommand1", Form1)
  With ctlCommand
    .Move 3450, 1, 450, 300
    .Caption = "Go!"
    .Visible = True
  End With
End Sub
```

