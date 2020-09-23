<div align="center">

## Using dir\(\) to get a list of files and directories


</div>

### Description

instead of using the api and having to deal with nulls and UDT's and stuff, y not just use dir(), i have included 2 functions that return string arrays which contain all the files or directorys in the folder u specify. enjoy :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mud Blud](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mud-blud.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mud-blud-using-dir-to-get-a-list-of-files-and-directories__1-32653/archive/master.zip)





### Source Code

```
Public Function GetFolderList(Path As String) As String()
Dim Dirs() As String, Cnt As Integer
Dim I As String
I = Dir$(Path, vbDirectory)
Do While I <> ""
If (GetAttr(Path & I) And vbDirectory) = vbDirectory Then
  If Trim$(I) = "." Or Trim$(I) = ".." Then GoTo DontAddItem
    If Cnt = 0 Then ReDim Dirs(0) Else ReDim Preserve Dirs(0 To Cnt + 1)
    Dirs(Cnt) = Path & Trim$(I)
    Cnt = Cnt + 1
End If
DontAddItem:
I = Dir$()
Loop
GetFolderList = Dirs()
End Function
Public Function GetFileList(Path As String, Match As String) As String()
Dim Files() As String, Cnt As Integer
Dim I As String
I = Dir$(Path & Match)
Do While I <> ""
  If Cnt = 0 Then ReDim Files(0) Else ReDim Preserve Files(0 To Cnt + 1)
  Files(Cnt) = Path & Trim$(I)
  Cnt = Cnt + 1
  I = Dir$()
Loop
GetFileList = Files()
End Function
```

