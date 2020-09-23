<div align="center">

## Get Free Resources


</div>

### Description

Gets Free Resources WITHOUT using a class module or a third party DLL as someone used below...
 
### More Info
 
Free Resources


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[i](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/i.md)
**Level**          |Advanced
**User Rating**    |3.0 (12 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/i-get-free-resources__1-7032/archive/master.zip)

### API Declarations

```
Public Declare Function pBGetFreeSystemResources Lib "rsrc32.dll" Alias "_MyGetFreeSystemResources32@4" (ByVal iResType As Integer) As Integer
```


### Source Code

```
Public Function SystemResources() As String
GDI$ = CStr(pBGetFreeSystemResources(1))
Sys$ = CStr(pBGetFreeSystemResources(0))
User$ = CStr(pBGetFreeSystemResources(2))
SystemResources$ = "GDI: " + GDI$ + "%"
SystemResources$ = SystemResources$ + vbCrLf + "System: " + Sys$ + "%"
SystemResources$ = SystemResources$ + vbCrLf + "User: " + User$ + "%"
End Function
'--------------------
'To use this code in a Message Box, use:
MsgBox SystemResources$, vbSystemModal, "System Resources"
'--------------------
'To use this code in a Text Box, use:
Text1 = SystemResources$
'Text1 being your Text Box name
'--------------------
'The SystemResources function was made to be placed in a module; if you would like it to be placed in your form... copy the declaration and function, paste it in your form coding, and change the Public to Private.
```

