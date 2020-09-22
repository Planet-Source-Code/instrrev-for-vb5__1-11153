<div align="center">

## InStrRev for VB5


</div>

### Description

This is a InStrRev function for VB5. I took a look at the one microsoft recomend, and almost died of laughter.
 
### More Info
 
The string to Search

The string to Find

Optional :> The start position

The postion of the Found string in the Searched string.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |3.0 (9 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/instrrev-for-vb5__1-11153/archive/master.zip)





### Source Code

```
Function myInStrRev(strStringToSearch As String, strFind As String, Optional iStart As Long) As Long
 Dim ip1 As Long, ip2 As Long
 Dim iLenStringToSearch As Long
 'get the length of the string
 iLenStringToSearch = Len(strStringToSearch)
 'if the start is 0 then set the start to the length
 'og the string
 If iStart = 0 Then
 iStart = iLenStringToSearch
 End If
 ip1 = 1
 Do
 ip2 = InStr(ip1, strStringToSearch, strFind)
 If (ip2 > 0) And (ip2 < iStart) Then
 'if ip2 is not zero and it is less than the
 'place to start searching then set the function
 'to return that position
 myInStrRev = ip2
 ElseIf ip2 = 0 Then
 ip2 = iLenStringToSearch
 End If
 'set the next position to seracf from
 ip1 = ip2 + 1
 Loop Until ip1 >= iStart
End Function
```

