<div align="center">

## Use of the Toolbar Control


</div>

### Description

Short tutorial on how to use the toolbar in VB4/5 32 bit. http://137.56.41.168:2080/VisualBasicSource/vb4toolbar.txt
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-use-of-the-toolbar-control__1-695/archive/master.zip)





### Source Code

```
'first be sure you have add the custom controls
'Microsoft Windows Common Controls -> Comctl32.ocx
'step 1:
'add the control Imagelist
'set the propertie Custom
'General: size
'Images: click 'Insert Picture' to add the necessary pictures
'step 2:
'add the control Toolbar
'set the propertie Custom
'General: Imagelist
'Buttons: to add a button just click on 'Insert Button'
'at 'Image' you need to set the index-number of the wanted picture
'this number is the same as the pictures index in the ImageList
'place - if you want - a ToolTipText
'or if you just want text place it behind the propertie 'Caption'
'click on 'OKE' when you are finished
'and the toolbar is ready
'now the code
'put it under the
Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
  Select Case Button.Index
  Case 1	'click on the first button
  Case 2	'click on the second button
  Case 3	'click on the third button
  	'and so on
  End Select
End Sub
'you can change most properties a runtime
Toolbar1.Buttons(1).Visible = False 'makes the first button disappear
Toolbar1.Buttons(1).ToolTipText = "an other one" 'change the tooltip text of the first button
Toolbar1.Buttons(2).Enabled = False 'disable the second button
Toolbar1.Buttons(3).Caption = "KATHER" 'change the caption of the third button
'BTW you cannot set the property Toolbar1.ShowTips at runtime!
```

