<div align="center">

## Strip HTML


</div>

### Description

Strip HTML from a string with this function. That basically explains it's all!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[DoWnLoHo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/downloho.md)
**Level**          |Unknown
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/downloho-strip-html__1-2994/archive/master.zip)





### Source Code

```
Public Function StripHTML(sHTML As String) As String
Dim sTemp As String, lSpot1 As Long, lSpot2 As Long, lSpot3 As Long
sTemp$ = sHTML$
Do
 lSpot1& = InStr(lSpot3& + 1, sTemp$, "<")
 lSpot2& = InStr(lSpot1& + 1, sTemp$, ">")
  If lSpot1& = lSpot3& Or lSpot1& < 1 Then Exit Do
  If lSpot2& < lSpot1& Then lSpot2& = lSpot1& + 1
 sTemp$ = Left$(sTemp$, lSpot1& - 1) + Right$(sTemp$, Len(sTemp$) - lSpot2&)
 lSpot3& = lSpot1& - 1
Loop
StripHTML$ = sTemp$
End Function
```

