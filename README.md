<div align="center">

## FormWinRegPos


</div>

### Description

This Procedure can be used by AnyForm to Get or Save the Form Position from the Windows Registry using SaveSetting and GetSetting :)
 
### More Info
 
pMyForm As Form

Optional pbSave As Boolean

Best to use this in either Form_Load, Form_Unload or Form_QueryUnload

Form_Load For Getting the Saved Form Posn Settings

Unload or QueryUnload for saveing Current Form Posn.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brad Skidmore](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brad-skidmore.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brad-skidmore-formwinregpos__1-6403/archive/master.zip)





### Source Code

```
Option Explicit
Public Sub FormWinRegPos(pMyForm As Form, Optional pbSave As Boolean)
  'This Procedure will Either Retrieve or Save Form Posn values
  'Best used on Form Load and Unload or QueryUnLoad
  On Error GoTo EH
  With pMyForm
    If pbSave Then
      'If Saving then do this...
      'If Form was minimized or Maximized then Closed Need to Save Windowstate
      'THEN... set Back to Normal Or previous non Max or Min State then Save
      'Posn Parameters SaveSetting App.EXEName, .Name, "Top", .Top
      SaveSetting App.EXEName, .Name, "WindowState", .WindowState
      If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
        .WindowState = vbNormal
      End If
      'Save AppName...FrmName...KeyName...Value
      SaveSetting App.EXEName, .Name, "Top", .Top
      SaveSetting App.EXEName, .Name, "Left", .Left
      SaveSetting App.EXEName, .Name, "Height", .Height
      SaveSetting App.EXEName, .Name, "Width", .Width
    Else
      'If Not Saveing Must Be Getting ..
      'Need to ref AppName...FrmName...KeyName (If nothing Stored Use The Exisiting Form value)
      .Top = GetSetting(App.EXEName, .Name, "Top", .Top)
      .Left = GetSetting(App.EXEName, .Name, "Left", .Left)
      .Height = GetSetting(App.EXEName, Name, "Height", .Height)
      .Width = GetSetting(App.EXEName, .Name, "Width", .Width)
      'Be Sure WindowState is set last (Can't Change POSN if vbMinimized Or Maximized
      .WindowState = GetSetting(App.EXEName, .Name, "WindowState", .WindowState)
    End If
  End With
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub
Private Sub Form_Load()
  FormWinRegPos Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  FormWinRegPos Me, True
End Sub
```

