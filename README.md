<div align="center">

## Another \(very easy\) Rounding Function


</div>

### Description

This code makes easy work of doing simple rounding on decimal number...something that VB currently lacks. Why didn't they think of this?
 
### More Info
 
Number (double data type)

RoundNum (Integer)

Rounding financial stuff is not cool :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-roberts.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-roberts-another-very-easy-rounding-function__1-6563/archive/master.zip)





### Source Code

```
' Paste this code directly into your IDE. This has not yet been tested on VBScript, but should work if you drop the type declarations.
Function RoundNum(Number As Double) As Integer
If Int(Number + 0.5) > Int(Number) Then
  RoundNum = Int(Number) + 1
Else
  RoundNum = Int(Number)
End If
End Function
```

