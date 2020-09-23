<div align="center">

## Retuning multiple parameters from functions


</div>

### Description

VB provides a very easy way in which to pass multiple parameters to subroutines and functions.

Whilst it is possible to return the results of processing in the passed parameters it is not very good practice, but many programmers do it anyway because they believe that VB functions will only return one parameter.

This simple example shows a clean method of returning as many parameters as you like from a function without resorting to modifying the passed parameters.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Phobos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/phobos.md)
**Level**          |Beginner
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/phobos-retuning-multiple-parameters-from-functions__1-31737/archive/master.zip)





### Source Code

```
Option Explicit
Type ReturnedParameters
  Parameter1 As String
  Parameter2 As Integer
  Parameter3 As Boolean
End Type
Private Sub main()
  ' Simple test program which shows how to return multiple parameters
  ' from a function.
  With TestFunction
    MsgBox .Parameter1
    MsgBox .Parameter2
    MsgBox .Parameter3
  End With
End Sub
Private Function TestFunction() As ReturnedParameters
  ' Example function showing how multiple parameters can be returned
  Dim sString As String, iInteger As Integer, bBoolean As Boolean
  sString = "Test String"
  iInteger = 12345
  bBoolean = True
  With TestFunction
    .Parameter1 = sString
    .Parameter2 = iInteger
    .Parameter3 = bBoolean
  End With
End Function
```

