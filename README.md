<div align="center">

## Shadow for controls


</div>

### Description

i saw many people creating shadow by adding a label control it dosent matter if you have min no of controls but when you have large no of controls

So here is the code to create shadow without using any other controls

Just paste this code ***Have Fun***
 
### More Info
 
Create shadow just using code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joseph Varghese](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joseph-varghese.md)
**Level**          |Beginner
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joseph-varghese-shadow-for-controls__1-39921/archive/master.zip)





### Source Code

```
Public Sub SHADOW(Frm As Form, Ctrl As Control)
Const SC = vbBlack 'you can define your own color
Const SW = 3
Dim IOLDWIDTH As Integer
Dim IOLDSCALE As Integer
IOLDWIDTH = Frm.DrawWidth
IOLDSCALE = Frm.ScaleMode
Frm.ScaleMode = 3
Frm.DrawWidth = 1
Frm.Line (Ctrl.Left + SW, Ctrl.Top + SW)-Step(Ctrl.Width - 1, Ctrl.Height - 1), SC, BF
Frm.DrawWidth = IOLDWIDTH
Frm.ScaleMode = IOLDSCALE
End Sub
Private Sub Form_paint()
SHADOW Me, Command1
End Sub
```

