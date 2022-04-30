<!--#include virtual="/include/CCHeader.asp"-->

<%
'another Gus Mueller automaton
'January 1999
fresh=1
ye1=cstr(year(date))
mo1=cstr(month(date))
if len(mo1)=1 then 
mo1 = "0" & mo1
end if
da1=cstr(day(date))
if len(da)=1 then 
da1 = "0" & da1
end if
set fsc=server.createobject("scripting.filesystemobject")
description=request("description")
title=request("title")
articlefile=request("articlefile")
peoplefile=request("peoplefile")
question=request("question")
changedate=request("changedate")
articlefilebackup=request("taco")
writer=request("writer")
if peoplefile="" and changedate="" and description<>"" then %>
You must select a person!
<%
response.end
end if
%>
<html>
<head>
<title>About Us Tool</title>
</head>
<body bgcolor=ccccff text=black>
<font face=arial>

<font face=arial>
<font size=-1>
<a href=cop.asp>copy robot</a> |
<a href=getdoc.asp>global campus editing</a> |
<a href=chatsched.asp>host editors</a> |
<a href=todaymaker.asp>today in cc</a> |
<a href=lovemaker.asp>Dr. Lovejoy</a> |
<a href=sexrated.asp>SexRated</a> |
<a href=styler.asp>StyleExperts</a>
People Editor
</font>
<br>
<font size=-1 face=arial>
Article |
<a href=sexerq.asp?peoplefile=<%=peoplefile%>>Question Entry</a> |
<a href=sexdescriptioncat.asp?peoplefile=<%=peoplefile%>>descriptions Catalog Updater</a> |
<a href=sexquestioncat.asp?peoplefile=<%=peoplefile%>>Question Catalog Updater</a>
</font>
<h2>About Us Article Tool: <% if articlefile<>"" then response.write(left(articlefile,4)&"-"&mid(articlefile,5,2)&"-"&right(articlefile,2))&" "&peoplefile%></h2>
Use this form to edit future About Us articles.
<form method="post" action="sexer.asp" name=beer>

<b>question:</b><br><font size=1>
<select name=peoplefile>
<option value="">--none--
<%
if fsc.fileexists(server.mappath("\aboutus\people\people\peoplecatalog.htm")) then
set peoplecatalog=fsc.opentextfile(server.mappath("\aboutus\people\people\peoplecatalog.htm"), 1, true)
if not peoplecatalog.atendofstream then
preperson=peoplecatalog.readall
end if
peoplecatalog.close
preperson=replace(preperson, chr(13), "")
preperson=trim(replace(preperson, chr(10), ""))
if right(preperson, 1)=chr(34) then
preperson=left(preperson, len(preperson)-1)
end if
if right(preperson,1)="|" then
preperson=left(preperson, len(preperson)-1)
end if
persons=split(preperson, "|")
for t=0 to ubound(persons)
persons(t)=replace(persons(t), chr(13), "")
persons(t)=replace(persons(t), chr(10), "")
if fsc.fileexists(server.mappath("\abouts\people\people\"&persons(t)&".htm")) then
'now we open the peoplefile and get the person
set persohn=fsc.opentextfile(server.mappath("\aboutus\people\people\"&persons(t)&".htm"), 1, true)
getums2=persohn.readall
question.close
getums2=replace(getums2, chr(13), "")
getums2=trim(replace(getums2, chr(10), ""))
if right(getums2, 1)=chr(34) then
getums2=left(getums2, len(getums2)-1)
end if
unpos=instr(getums2, "<$un>")
username=left(getums2, unpos-1)

end if
%>
<option value="<%=persons(t)%>" <%if peoplefile=persons(t) then%>selected<%end if%>><%=username%>
<%

next
end if 'catalog file exists
%>
</select>&nbsp;<a href=peopleq.asp?changedate=EE&taco=<%=peoplefile%>&peoplefile=<%=peoplefile%>>edit</a>
</font><br>
<%=hoolah%>

<p> 
<%
if changedate="" then 
moder="write"
else
moder="changedate"
end if 
if articlefile="" then 
articlefile=ye1&mo1&da1
else
if description<>"" then

if moder="write" and articlefilebackup=articlefile then
entry=title & "<$tq>" & question & "<$dq>" & description

set fullarticlefile=fsc.createtextfile(server.mappath("\aboutus\people\descriptions\"&articlefile&peoplefile&".htm"), true)
fullarticlefile.writeline(entry)
fullarticlefile.close
response.write("<table bgcolor=ffffff width=400><td><font color=cc00cc face=arial><h2>"&title&"</h2><font color=660066>"&question&"</font><P>"&description&"<p></font></table><br>was written to Dev.</font>")

end if 'moder=write
end if
end if 'articlefile

if moder="changedate" and peoplefile<>"" then

if fsc.fileexists(server.mappath("\aboutus\people\articles\"&articlefile&peoplefile&".htm")) then
set fullarticlefile=fsc.opentextfile(server.mappath("\aboutus\people\articles\"&articlefile&peoplefile&".htm"), 1)
getums=trim(fullarticlefile.readall)
fullarticlefile.close

'parse out chunks
arb=instr(getums, "<$tq>")
qub=instr(getums, "<$qd>")

lenny=len(getums)
title=cleanup(left(getums, arb-1))
question=cleanup(mid(getums, arb+5, (qub-arb-5)))
description=trim(mid(getums, qub+5, (lenny-qub)))

%>
<b>title:</b>
<input name=title type=text value="<%=title%>">
<p><b>description</b><br>
<TEXTAREA NAME="description" ROWS="6" COLS="55" align=middle> 
<%=question%></textarea>
<TEXTAREA NAME="description" ROWS="6" COLS="55" align=middle> 
<%=description%></textarea><br>
<%else%>
<br>
<font color=red>
note: this entry didn't exist.</font><br>
<b>title:</b>
<input name=title type=text value="">
<p><b>description</b><br>
<TEXTAREA NAME="description" ROWS="6" COLS="55" align=middle></textarea>
<TEXTAREA NAME="description" ROWS="6" COLS="55" align=middle></textarea><br>
<%end if%><br>
<%end if
%><p>
<%if peoplefile<>"" then%>
<select name=articlefile onChange="document.beer.changedate.click();">
<%
for daze=-30 to 21
datem=dateserial(year(date), month(date), day(date)+(daze))
da=day(datem)
mo=month(datem)
ye=year(datem)
if len(mo)=1 then 
mo = "0" & mo
end if
if len(da)=1 then 
da = "0" & da
end if
if ye&mo&da=articlefile then
sel="selected"
else
sel=""
end if

'look to see if there is a legitimate file here
ex=""
if fsc.fileexists(server.mappath("\aboutus\people\descriptions\"&ye&mo&da&peoplefile&".htm")) then
set ourfile=fsc.getfile(server.mappath("\aboutus\people\descriptions\"&ye&mo&da&peoplefile&".htm"))
ourfilesize=ourfile.size
if ourfilesize>4 then 
ex="*"
end if
end if


%>
<option <%=sel%> value=<%=ye&mo&da%>><%=mo&"/"&da&"/"&ye&ex%>
<%next%>
</select> day&nbsp;
<%
buttoner="Select Article Date"
else 'question file exists
buttoner="Select Question Type"
end if
%>
<INPUT TYPE=SUBMIT value="<%=buttoner%>" name="changedate">
<%if moder<>"write" and peoplefile<>"" then%>
<INPUT TYPE=SUBMIT value="Write to Dev" name="writer">
<%end  if%>
<INPUT TYPE=hidden value="<%=articlefile%>" name="taco">
</form>
<p><%
if moder="write" and articlefilebackup=articlefile then

if not fsc.fileexists(server.mappath("\trans\"&mo1&da1&"people.tra")) then
set transcript=fsc.createtextfile(server.mappath("\trans\"&mo1&da1& "people.tra"), true)
transcript.close
end if


set transcript=fsc.opentextfile(server.mappath("\trans\"&mo1&da1& "people.tra"), 8)
transcript.writeline("\ccroot\aboutus\people\descriptions\"&articlefile&peoplefile&".htm")
transcript.close
response.write("<form method=post action=http://content/update/copy.asp><font size=+1 color=red>People articles have been modified</font><br>")
response.write("SexRated file:<input type=text name=schedfile value="&mo1&da1&"people.tra><INPUT TYPE=SUBMIT value="&chr(34)&"Copy the files to the live server"&chr(34)&"><input type=hidden name=username value="&chr(34)&"sexrated front end"&chr(34)&"></form>")

end if 'there should be a transcript update
%>
<ul>
<font face=arial>
<p>
When you're done making a batch of People articles, you should<br>
<b><a href=sexdescriptioncat.asp?peoplefile=<%=peoplefile%>>Update</a> the People description Catalog.</b><p>
Go check out the <a href=/aboutus/people/people.asp?peoplefile=<%=peoplefile%>&articlefile=<%=articlefile%>>People front end</a>.

<p></ul>
<font size=-1 face=arial>
Article Entry |
<a href=people.asp?peoplefile=<%=peoplefile%>>Question Entry</a> |
<a href=sexdescriptioncat.asp?peoplefile=<%=peoplefile%>>descriptions Catalog Updater</a> |
<a href=sexquestioncat.asp?peoplefile=<%=peoplefile%>>Question Catalog Updater</a>
</font>
</body>
</html>

<%

function cleanup(inner)
cleanup=trim(replace(trim(inner), chr(13), ""))
cleanup=replace(cleanup, chr(10), "")
end function

%>

<!--#include virtual="/include/CCFooter.asp"-->
