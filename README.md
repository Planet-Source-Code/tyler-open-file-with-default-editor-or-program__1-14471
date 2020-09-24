<div align="center">

## Open File With Default Editor or Program


</div>

### Description

This 1 line of code will Open any file with their correct default program. EX: opens .Doc file with winword and .Zip file with WinZip..Or any file with their default program..
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tyler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tyler.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tyler-open-file-with-default-editor-or-program__1-14471/archive/master.zip)

### API Declarations

```
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
```


### Source Code

```
Private Function openfile(file As String)
Call ShellExecute(0&, vbNullString, file, vbNullString, vbNullString, vbNormalFocus)
End Function
'' That's it!..
'' Please Vote..:-)
```

