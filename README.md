<div align="center">

## Check for extended ASCII


</div>

### Description

Checks to make sure the contents of a TextBox or String are only standard ASCII; If it has any extended ASCII, the function will return False.
 
### More Info
 
Use like this:

If Valid_ASCII(Text1) = False Then

MsgBox "Invalid chars!"

Exit sub

End If


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alex Bylund](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex-bylund.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alex-bylund-check-for-extended-ascii__1-30010/archive/master.zip)





### Source Code

```
Function Valid_ASCII(Text As String)
 For x = 1 To Len(Text)
  If Asc(Mid(Text, x, 1)) > 126 Or Asc(Mid(Text, x, 1)) < 32 Then
   Valid_ASCII = False
   Exit Function
  End If
 Next x
 Valid_ASCII = True
End Function
```

