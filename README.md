<div align="center">

## Unique w/ database & more than one page


</div>

### Description

This is an update to my Unique Hit Counter that

used a text file to keep track of all hits. With

a database, more than one page on your site can be

monitored by this file. It uses the servervariable

url to determine what page it needs to add a hit

to.

As with my last hit counter, it uses cookies to

keep track of who has been and who hasn't. I built

my Text to Images into this script. As with that

submission, you have to make your own images. They

don't have to be any special height, or width. My

images start with "cnt_" and have a numerical

value coresponding with the value passed (that

means you need images cnt_0 - cnt_9). I also have

a cnt_start and cnt_end image to make everything

look nice.

Save this save as something like counter.asp and

just use an include (<!-- #include file="counter.asp" //-->). You can use it on every

page on your site to track where people go, and

what-not.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[atwinda](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atwinda.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atwinda-unique-w-database-more-than-one-page__4-6628/archive/master.zip)





### Source Code

```
<%
Dim objHitConn, strHitSQL, objHitRs, intHits, strHitFile
Set objHitConn = Server.CreateObject("ADODB.Connection")
objHitConn.Provider = "Microsoft.Jet.OLEDB.4.0"
objHitConn.Open Server.MapPath("counter.mdb")
'counter.mdb needs to have a table named "Main"
'along with two colunms: "Page" and "Hits"
strHitFile = Request.ServerVariables("url")
strHitSQL = "SELECT Page, Hits From Main Where Page='" & strHitFile & "'"
Set objHitRs = Server.CreateObject("ADODB.Recordset")
objHitRs.Open strHitSQL, objHitConn, 1, 2
If objHitRs.EOF Then
	objHitRs.AddNew
	intHits = 0
	objHitRs.Fields("Page").Value = strHitFile
Else
	intHits = objHitRs.Fields("Hits").Value
End If
intHits = CInt(intHits) + 1
objHitRs.Fields("Hits").Value = CStr(intHits)
objHitRs.Update
objHitRs.Close
objHitConn.Close
set objHitRs = nothing
set objHitConn = nothing
Call DisplayImg(intHits)
Function DisplayImg(intNum)
Dim itmCur, tmpCur
Response.Write "<img src='images/cnt_start.gif'>"
For itmCur = 1 To Len(intNum)
	tmpCur = Mid(cStr(intNum), itmCur, 1)
	Response.Write "<img src='images/cnt_" & tmpCur & ".gif'>"
Next
Response.Write "<img src='images/cnt_end.gif'>"
End Function
%>
```

