<div align="center">

## Example of how to auto scroll


</div>

### Description

Teaches you how to auto-scroll a richtextbox, but it accounts for the vertical scrollbar. If the user has the scrollbar moved up, or at the top, it does not scroll down and stay at the bottom. This is very usefull for logging windows or chat windows. There is no code to do this on PSC. I would like to thank BillSoo from www.visualbasicforum.com for his help in solving this. He gave me the initial code and sent me in the right direction. After fixing it, here it is for everyone to use. This is just a little demo, the timer runs every second adding a line to the box, just watch it go and then play with the scrollbar, it only scrolls down if its already at the bottom.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-05-14 18:24:42
**By**             |[Gregg Housh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gregg-housh.md)
**Level**          |Intermediate
**User Rating**    |4.2 (46 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Example\_of831335142002\.zip](https://github.com/Planet-Source-Code/gregg-housh-example-of-how-to-auto-scroll__1-34792/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
Private Const EM_GETSCROLLPOS = (WM_USER + 221)
Private Const EM_SETSCROLLPOS = (WM_USER + 222)
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_CHARFROMPOS = &HD7
Private Const EM_GETLINECOUNT = &HBA
Private Type POINTL
x As Long
y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
```





