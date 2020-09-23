<div align="center">

## How to round a byte amount into KB or MB \(also KB/s\) with two decimal places


</div>

### Description

Apart from what it says on the title, I will also demostrate how to calculate, based on a byte amount and a time period, the amount of KB/s (Kilobytes per second) of a transfer. The decimal separator (. or ,) is shown according to the regional settings.
 
### More Info
 
REMEMBER: It is good programming practice to declare your variables AND to make functions out of routines so that you can reuse or modify them easily in the future, this can also be applied to a group of functions and subs that should do something together, just put them in a module or class. Also, do not forget to always comment your code!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Luis Cantero](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/luis-cantero.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/luis-cantero-how-to-round-a-byte-amount-into-kb-or-mb-also-kb-s-with-two-decimal-places__1-38018/archive/master.zip)





### Source Code

```
'PURPOSE:  Rounds a Byte amount and returns KB with 2 decimal places
'INPUT:   Long: Byte amount
'OUTPUT:  String: Rounded KB amount
Function GetRoundedKB(lngByteAmount As Long) As String
  GetRoundedKB = FormatNumber(Int(lngByteAmount / 1024 * 100 + 0.5) / 100, 2)
End Function
'PURPOSE:  Rounds a Byte amount and returns, according to an elapsed time in seconds, KB/s with 2 decimal places
'INPUT:   Long: Byte amount
'OUTPUT:  String: Rounded KB/s amount
Public Function GetRoundedKBperS(lngByteAmount As Long, lngSecondsElapsed As Double) As String
  'Error check
  If lngSecondsElapsed <= 0 Then lngSecondsElapsed = 1
  GetRoundedKBperS = FormatNumber(Int(lngByteAmount / 1024 / lngSecondsElapsed * 100 + 0.5) / 100, 2)
End Function
'PURPOSE:  Rounds a Byte amount and returns MB with 2 decimal places
'INPUT:   Long: Byte amount
'OUTPUT:  String: Rounded MB amount
Public Function GetRoundedMB(lngByteAmount As Long) As String
  GetRoundedMB = FormatNumber(Int(lngByteAmount / 1048576 * 100 + 0.5) / 100, 2)
End Function
'Here's sample source code for an API that rounds a byte amount,
'In my opinion, it is just too much for too little...:
Private Declare Function StrFormatByteSize Lib _
	"shlwapi" Alias "StrFormatByteSizeA" (ByVal _
	dw As Long, ByVal pszBuf As String, ByRef _
	cchBuf As Long) As String
Public Function FormatKB(ByVal Amount As Long) _
	As String
	Dim Buffer As String
	Dim Result As String
	Buffer = Space$(255)
	Result = StrFormatByteSize(Amount, Buffer, _
		Len(Buffer))
	If InStr(Result, vbNullChar) > 1 Then
		FormatKB = Left$(Result, InStr(Result, _
			vbNullChar) - 1)
	End If
End Function
```

