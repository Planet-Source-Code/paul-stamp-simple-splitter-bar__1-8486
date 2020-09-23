<div align="center">

## Simple Splitter Bar


</div>

### Description

The simplest (as few as three lines of code needed!) and best code to implement explorer-style splitter bars
 
### More Info
 
This sample takes advantage of two seldom used features of picture boxes - the Align property and the Resize event. I first discovered this technique when I had an MDI application that was to have a treeview control on the left of the screen, but the right half of the screen should be the workspace where child windows appear. The only controls that can be added directly onto an MDI parent form are those with Align properties - i.e. controls that will snap to the bottom, top, left or right of the screen. Unfortunately the treeview control does not support this, so I added a picture box called "picPane" and set it's Align property to 3 (Align Left). Then I put my treeview onto this picture box - and scratched my head as to how to get it to resize

At first I tried some of the solutions put forward on various tips websites, but their solutions were fairly inelegant so I experimented a little bit and hit upon this :

Add a new picture box, call it "picSplitter" and set Align property to 3 (AlignLeft), and make the width 60 twips.

Set the mousepointer property of the picturebox to 9 (Size WE)

Add the following code to the mousemove event of the picturebox :

If button = vbLeftButton then

picPane.Width = picSplitter.left + X

End If

This will then fire picPane's Resize event, in response to which you resize your controls on picPane

Simple as that !

Of course its not quite as simple as that, but that is the principle. It works because if you move the pointer outside of the boundary of the picture box but have the mouse button pressed, the X property will go negative. Use the X value as an offset from the left position of the splitter bar and because the splitter will "snap" against the picPane picturebox, setting the width of picPane will have the effect of moving the bar as well as widening/narrowing the pane. Trust me.. it works, and not just for MDI forms... it will work for any form, for more than one pane (ie have a treeview and listview, like in Explorer), and for horizontal bars as well as vertical bars.


<span>             |<span>
---                |---
**Submitted On**   |2000-03-26 16:03:54
**By**             |[Paul Stamp](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-stamp.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD62535312000\.zip](https://github.com/Planet-Source-Code/paul-stamp-simple-splitter-bar__1-8486/archive/master.zip)








