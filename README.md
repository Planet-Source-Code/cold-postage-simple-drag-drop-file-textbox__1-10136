<div align="center">

## simple drag drop file textbox


</div>

### Description

drag and drop files into a textbox is simple with this code
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[cold\_postage](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cold-postage.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 6\.0, VB Script
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cold-postage-simple-drag-drop-file-textbox__1-10136/archive/master.zip)





### Source Code

create a textbox on an empty form<br>
in the property window of the textbox change the OLEDropMode to "Manual".
<br>
<b>now add this function to your form code:</b>
<br>
<br>
Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
<br>
<br>
 If Data.GetFormat(vbCFFiles) Then Text1.Text = Data.Files(1)
<br>
<br>
End Sub
<br>
<b>add the following if you don't want to show the drag drop mouse pointer when the item is not a file </b>
<br>
Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)<br><br>
 If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
<br>
<br>
End Sub

