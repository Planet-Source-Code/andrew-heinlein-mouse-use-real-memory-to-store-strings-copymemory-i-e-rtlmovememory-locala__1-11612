<div align="center">

## Use REAL memory to store strings\. \(CopyMemory i\.e RtlMoveMemory, LocalAlloc, LocalFree API\)


</div>

### Description

Use your machines real memory to store large strings instead of varibles that run down your programs resources.
 
### More Info
 
fun stuff

because you are allocating real memory, your program may crash in DEBUG mode. it's rare though.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew Heinlein \(Mouse\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-heinlein-mouse.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-heinlein-mouse-use-real-memory-to-store-strings-copymemory-i-e-rtlmovememory-locala__1-11612/archive/master.zip)

### API Declarations

```
'Put this in a MODULE
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const LPTR = (&H0 Or &H40)
```


### Source Code

```
'Put this code in the SAME MODULE as the API ABOVE
'if you would like to download a working example this code go here:
'http://www.theblackhand.net/mouse/RealMemory.zip
Public Function malloc(Strin As String) As Long
 Dim PointerA As Long, lSize As Long
 lSize = LenB(Strin) 'Length of string in bytes.
 'Allocate the memory needed and returns a pointer to that memory
 PointerA = LocalAlloc(LPTR, lSize + 4)
 If PointerA <> 0 Then
  'Final allocation
  CopyMemory ByVal PointerA, lSize, 4
  If lSize > 0 Then
   'copy the string to that allocated memory.
   CopyMemory ByVal PointerA + 4, ByVal StrPtr(Strin), lSize
  End If
 End If
 'return the pointer to the string stored memory
 malloc = PointerA
End Function
Public Function RetMemory(PointerA As Long) As String
 Dim lSize As Long, sThis As String
 If PointerA = 0 Then
  GetMemory = ""
 Else
  'get the size of the string stored at pointer "PointerA"
  CopyMemory lSize, ByVal PointerA, 4
  If lSize > 0 Then
   'buffer a varible
   sThis = String(lSize \ 2, 0)
   'retrive the data at the address of "PointerA"
   CopyMemory ByVal StrPtr(sThis), ByVal PointerA + 4, lSize
   'return the buffer
   RetMemory = sThis
  End If
 End If
End Function
Public Sub FreeMemory(PointerA As Long)
 'frees up the memory at the address of "PointerA"
 LocalFree PointerA
End Sub
```

