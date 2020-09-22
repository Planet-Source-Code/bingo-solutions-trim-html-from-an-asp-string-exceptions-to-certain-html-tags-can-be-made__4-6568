<div align="center">

## Trim HTML from an ASP String \(exceptions to certain HTML tags can be made\)


</div>

### Description

Use the TrimHTML function to remove any HTML strings from the text. Also, when you can specify certain tags which you DONT want it to remove. This is especially useful when you want to allow your users to be able to specify certain tags like <B> or <I>. It is simple function which can save great deal of time. At first, I like the prebuilt HTMLEncode function provided with ASP but again, it removed all HTML tags and did not give opputunity for us to make exceptions to certain formatting tags like <B>. I hope you like it and please vote for me!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bingo Solutions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bingo-solutions.md)
**Level**          |Beginner
**User Rating**    |3.7 (26 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Security](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/security__4-14.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bingo-solutions-trim-html-from-an-asp-string-exceptions-to-certain-html-tags-can-be-made__4-6568/archive/master.zip)





### Source Code

```
INSTRUCTIONS:
Call the function from you code like below:
strResults=TrimHTML ("Plain Text.<BR><B>Should be bold!</B><BR><b>Should be bold</B><BR><i>should be italics</i><BR><B><I>Bold + italics!</B></i>","b,i,BR")
In this case, strRestuls will hold the new text. Notice that I have specified 3 tags which I dont want it to remove from the string. You can do as many as you want but remember that you should not put the "<" and ">" around them when calling this function.
If you want it to remove all HTML tags, then use then specify "" as the second perimeter when calling the function... heres an example:
strResults=TrimHTML ("Plain Text.<BR><B>Should be bold!</B><BR><b>Should be bold</B><BR><i>should be italics</i><BR><B><I>Bold + italics!</B></i>","")
---------------------------------------
<%
strResults=TrimHTML ("Plain Text.<BR><B>Should be bold!</B><BR><b>Should be bold</B><BR><i>should be italics</i><BR><B><I>Bold + italics!</B></i>","")
 Response.Write strResults
Function TrimHTML (strHTML,arrDontTrim)
 strHTML=Replace(strHTML,"<","&#60;")
strHTML=Replace(strHTML,">","&#62;")
If len(arrDontTrim)<>0 Then
arrDontTrim=Split(arrDontTrim,",")
For i = 0 to ubound(arrDontTrim)
strHTML=Replace(strHTML,"&#60;" & arrDontTrim(i) & "&#62;","<" & arrDontTrim(i) & ">",1, -1, vbTextCompare)
strHTML=Replace(strHTML,"&#60;/" & arrDontTrim(i) & "&#62;","</" & arrDontTrim(i) & ">",1, -1, vbTextCompare)
Next
End If
 TrimHTMl= strHTML
End Function
%>
```

