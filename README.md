<div align="center">

## Connect, Disconnect Networks Drives \( EASILY \)


</div>

### Description

Connects and Disconnects Network Drives from your System
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MASTER VB PROGRAMMER](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/master-vb-programmer.md)
**Level**          |Unknown
**User Rating**    |6.0 (605 globes from 101 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/master-vb-programmer-connect-disconnect-networks-drives-easily__1-2209/archive/master.zip)

### API Declarations

```
Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Public Const RESOURCETYPE_DISK = &H1
```


### Source Code

```
Private Sub cmdConnect_Click()
Dim x As Long
If Index = 0 Then
x = WNetConnectionDialog(Me.hwnd, RESOURCETYPE_DISK)
End If
End Sub
Private Sub cmdDisconnect_Click()
If Index = 1 Then
x = WNetDisconnectDialog(Me.hwnd, RESOURCETYPE_DISK)
End If
End Sub
```

