<div align="center">

## Possible error in VB6 Compiler


</div>

### Description

This is possebly an error in the VB6 compiler but i have turned it into a way to find out if you are in the IDE or Compiled version.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marco v/d Berg](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marco-v-d-berg.md)
**Level**          |Beginner
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marco-v-d-berg-possible-error-in-vb6-compiler__1-52220/archive/master.zip)





### Source Code

Just simple arithmatic can give different values in IDE-mode ore EXE-mode.<BR>
Insert this piece of code into a module and call the function IsIde.<BR> it will return true if the programm is running in IDE-mode and false if running in EXE-mode (Compiled).<BR>
Take a look at the code and find out if this is an error or not.<BR><BR>
<PRE>
Option Explicit
Private Test As Long
Public Function IsIde() As Boolean
 Test = 4
 Test = (Test + 1 + GetVal)
 If Test = 9 Then IsIde = True
'Test will give 9 if IDE and 10 if EXE
'This can be done with any initial value but ofcourse you get other results
End Function
Private Function GetVal() As Long
 GetVal = 4
 Test = Test + 1
End Function
</PRE>

