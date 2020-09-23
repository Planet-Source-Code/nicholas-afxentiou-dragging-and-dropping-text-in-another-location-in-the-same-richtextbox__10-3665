<div align="center">

## Dragging and Dropping Text in another location in the same richtextbox


</div>

### Description

For those who want to create their own text editor and want to move text around with drag and drop just like MS Word does.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nicholas Afxentiou](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nicholas-afxentiou.md)
**Level**          |Advanced
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB\.NET
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__10-26.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nicholas-afxentiou-dragging-and-dropping-text-in-another-location-in-the-same-richtextbox__10-3665/archive/master.zip)





### Source Code

'Add a richtextbox control in ur form<br<br>
<p>Dim sel1, sel2 as Integer <br>
<br>
Sub Richtextbox1_MouseDown(ByVal sender As System.Object, ByVal e As
System.Windows.Forms.MouseEventArgs) Handles Richtextbox1.MouseDown<br>
If Richtextbox1.SelectedText <> "" And e.Clicks <> 2 Then<br>
sel1 = Richtextbox1.SelectionStart<br>
sel2 = Richtextbox1.SelectionLength<br>
Richtextbox1.DoDragDrop(Richtextbox1.SelectedRtf, DragDropEffects.Move)<br>
End If<br>
End Sub<br>
<br>
<br>
Private Sub Richtextbox1_DragEnter(ByVal sender As Object, ByVal e As
System.Windows.Forms.DragEventArgs) Handles Richtextbox1.DragEnter<br>
e.Effect = DragDropEffects.Move<br>
End Sub<br>
<br>
Private Sub Richtextbox1_DragDrop(ByVal sender As Object, ByVal e As
System.Windows.Forms.DragEventArgs) Handles Richtextbox1.DragDrop<br>
If Richtextbox1.SelectionStart < sel1 Then<br>
Dim selStart As Int16 = Richtextbox1.SelectionStart<br>
Richtextbox1.SelectedRtf = e.Data.GetData(DataFormats.Text).ToString()<br>
Richtextbox1.SelectionStart = sel1 + sel2<br>
Richtextbox1.SelectionLength = sel2<br>
Richtextbox1.SelectedText = ""<br>
Richtextbox1.SelectionStart = selStart<br>
End If<br>
If Richtextbox1.SelectionStart > sel1 + sel2 Then<br>
Richtextbox1.SelectedRtf = e.Data.GetData(DataFormats.Text).ToString()<br>
Richtextbox1.SelectionStart = sel1<br>
Richtextbox1.SelectionLength = sel2<br>
Richtextbox1.SelectedText = ""<br>
End If<br>
End Sub</p>
<br><br>
 Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load<br>
    RichTextBox1.AllowDrop = True<br>
  End Sub

