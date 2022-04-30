<!--#include virtual="/include/CCHeader.asp"-->

<!-- include virtual="/include/default.asp"-->
<%
'Pod Builder front end
'code developed by Gus Mueller
'March 1999

'get the information about this pod

on error goto 0
tool=request("tool")
toolfolder="tools"
topmostdirectory="\extension"
toolpath=topmostdirectory & "\" & toolfolder
if tool="" then tool="faq"
set fso=server.createobject("scripting.filesystemobject")
	bgcolor=readlines(0)
	medium=readlines(1)
	text=readlines(2)
	link=readlines(3)
	vlink=readlines(4)
	'tool=readlines(5)
	strToolName=readlines(6)
	strShortPath=readlines(7)
	responseintro=readlines(8)
	imageheight=readlines(9)
	imagewidth=readlines(10)
	frontendurl=readlines(11)
	keyedto=readlines(12)
	fieldforkeyedpulldown=readlines(13)
	fieldforbigkeyedpulldown=readlines(14)
	keyname=readlines(15)
	topmostdirectory=readlines(16)
	bottommostdirectory=readlines(17)
	responseemail=readlines(18)
	fieldforpulldown=readlines(19)
	ftitle=readlines(20)
	introtext=readlines(21)
	thankyou=readlines(22)
	communityid=readlines(23)
	datefed=readlines(24)
	lister=readlines(25)
	brought=readlines(26)
	style=readlines(27)
	sectionid=readlines(28)
i=30
do until readlines(i)="$devoid$"
	i=i+1
loop
fieldnumber=i-31
i=30
redim fieldtype(fieldnumber)
for fields=0 to fieldnumber
	if fields<=ubound(fieldtype) then
	fieldtype(fields)=readlines(fields+30)
	end if
next

strTopPath=topmostdirectory & "\" & strShortPath
strKeyedPath=topmostdirectory & "\" & keyedto

white="ffffff"
'folder="love_relationships"
nowage=request("nowage")
questionfile=request("questionfile")
placement=request("placement")
keyplace=request("keyplace")
status=request("status")
FileDatePart=request("FileDatePart")


if placement<>"" then
	if not isnumeric(placement) then
		placement=cint(placement)
	end if
end if
if keyplace<>"" then
	if not isnumeric(keyplace) then
		keyplace=cint(keyplace)
	end if
end if

if placement="" and keyplace="" and status<>"thankyou" then status="new"

'get key catalog
if keyedto<>"" then
	set cata=fso.opentextfile(server.mappath(strKeyedPath & bottommostdirectory & "\catalog.htm"), 1, true)
	if not cata.atendofstream then
		catalog=trim(cata.readall)
	end if
	cata.close
	catalog=replace(catalog, chr(13), "")
	catalog=replace(catalog, chr(10), "")
	if right(catalog, 1)="|" then
		catalog=left(catalog, len(catalog)-1)
	end if
	arrKey=split(catalog, "|")
	qatop=ubound(arrKey)
	if keyplace=""  then 
		keyeddatepart=arrKey(qatop)
		keyplace=qatop
	else 
		questionfile=arrKey(keyplace)
	end if
end if

'get base catalog
set cata=fso.opentextfile(server.mappath(strTopPath & bottommostdirectory & "\" &  keyeddatepart & "catalog.htm"), 1, true)
if not cata.atendofstream then
	catalog=trim(cata.readall)
else
	nofile="<b>There are no responses yet for this "& keyname &".</b>"
	placement=0
end if
cata.close
if instr(catalog, "|")>0 then
	catalog=replace(catalog, chr(13), "")
	catalog=replace(catalog, chr(10), "")
	if right(catalog, 1)="|" then
		catalog=left(catalog, len(catalog)-1)
	end if
	arrFile=split(catalog, "|")
	intFileTop=ubound(arrFile)
else
	nofile="<b>There are no responses yet for this "& keyname &".</b>"
	placement=0
end if

if FileDatePart<>"" then
	placement=filetoplacement(FileDatePart)
end if

'calculate a possible article for today
if nofile="" then
	ye=cstr(year(date))
	mo=cstr(month(date))
	if len(mo)=1 then 
		mo = "0" & mo
	end if
	da=cstr(day(date))
	if len(da)=1 then 
		da = "0" & da
	end if
	pre=strTopPath & bottommostdirectory 
	pre2=strKeyedPath & bottommostdirectory 
	nowage=ye&mo&da
	if placement="" then
	
	for t=0 to intFileTop
		if makedate(arrFile(t)) <= date then
			FileDatePart=arrFile(t)
			placement=t
		end if
	next
	else
		FileDatePart=arrFile(placement)
	end if
	if placement="" then
		placement=intFileTop
		FileDatePart=arrFile(intFileTop)
	end if
end if


%>
<!-- include virtual="/campus/buttondetails/love_relationships.asp"-->
<html>
<head>
<title><%=ftitle%></title>

</head>

<body bgcolor=<%=bgcolor%> text=<%=text%> link=<%=medium%> vlink=<%=vlink%>>

<script language = "JavaScript">
<!--
function makeRemote() {
remote = window.open("","remotewin","height=290,width=300,scrollbars=yes");
remote.location.href = "profile.html";
    if (remote.opener == null) remote.opener = window;
remote.opener.name = "opener";
}
//-->
</script>

<table>
<tr>
<%
'left navbar
%>
<td valign="TOP" bgcolor=<%=medium%> width=<%if dateget="" then %>160<%else%>184<%end if%>><center><%
alsofeatured=replace(alsofeatured, "<A href=/libido/sexrated/ target=_top><img src=/ccgifs/campus/love_relationships/sexrated2.gif  height=29 border=0 vspace=4 ></a>", "")

response.write(alsofeatured)
%>
<img src=/ccgifs/travel/spacer.gif width=180 height=1>
<p>
<%if dateget<>"" then %>
<%end if%>
</td>
<%
'bulk of the column
%>
<td valign=top>
<table border="0" cellpadding="5" cellspacing="0"  width=393>
<tr bgcolor="<%=light%>">  
<td colspan=2><font size="1" color="<%=grey%>" face="Arial"> <font color="<%=medium%>">&nbsp;&nbsp;&nbsp; <%=brought%></font></font><font size="4" color="#000000" face="ARIAL, helvetica"> 
</font><font size="4" color="#004080" face="ARIAL, helvetica"> </font>
</tr>
<%if dateget="" then %>
	<tr bgcolor="<%=medium%>"> 
	<td>
	<center><nobr>
	<%if status<>"new" then
	if fso.fileexists(server.mappath("/ccgifs" & forwardslash(topmostdirectory) & "/" & tool &"/smallintro.gif")) then
%>
		<a href=sexrated.asp><img src=/ccgifs<%=forwardslash(topmostdirectory)%>/<%=tool%>/smallintro.gif width=178 height=57 alt="" border="0"></a>
	<%else%>
	<font color=<%=bgcolor%> size=6 face=arial,helvetica><b><%=ftitle%></b></font>
	<%end if%>	
		
		<%else
			if fso.fileexists(server.mappath("/ccgifs" & forwardslash(topmostdirectory) & "/" & tool &"/intro.gif")) then
		
		%>
		<img src=/ccgifs<%=forwardslash(topmostdirectory)%>/<%=tool%>/intro.jpg width=251 height=122 alt="" border="0"><%
else%>
	<font color=<%=bgcolor%> size=5 face=arial,helvetica><b><%=ftitle%></b></font>
	<%
end if
end if%>
		</center>
		<font face=verdana, helvetica size=1>
		</font></td></tr>
		<%	if keyedto<>"" then%>
		<tr>
		<td>
		
		<p><font face="Verdana, Arial, Helvetica, sans-serif">
		&nbsp;
		<%
		'get question
	
			set filKey=fso.opentextfile(server.mappath(strKeyedPath &"\"& keyeddatepart &".htm"), 1)
			question=trim(filKey.readall)
			filKey.close
		
		'at this point we need to parse out the pieces of the keyfile
		%>
		
		<h2>sample question</h2>
		
		<%
		

		%>
		</font>
		</td>
		</tr>
	<%end if
end if
'generate some kind of menu here
%>


<%
if status<>"old_new" then 
	%>
	<tr bgcolor="<%=light%>"> 
	<%if dateget="" then%>
		<td colspan=2> <font size="4" face="Arial, Helvetica, sans-serif" color="<%=medium%>">&nbsp;&nbsp;&nbsp;
		Latest <%=ftitle%><%if headersmall<>"" then%>: <%=headersmall%><%end if%><%
	else
		%><td colspan=2 bgcolor=<%=white%>> <font size="4" face="Arial, Helvetica, sans-serif" color="<%=medium%>">&nbsp;&nbsp;&nbsp;
	<%end if
	%></font></td>
	</tr><tr><td>
	<%
	'get answer
	if nofile="" then  
	'response.write("********" &strTopPath & bottommostdirectory & "\" & keyeddatepart & FileDatePart&".htm")
		if fso.fileexists(server.mappath(strTopPath & bottommostdirectory & "\" & keyeddatepart & FileDatePart&".htm")) then
			set fullFileDatePart=fso.opentextfile(server.mappath(strTopPath & bottommostdirectory & "\"& keyeddatepart & FileDatePart&".htm"), 1)
			getums=trim(fullFileDatePart.readall)
			fullFileDatePart.close
			
			'parse out and display fields of data, a major core part of this front end
			'response.write(getums)
			
			
						for t=0 to fieldnumber
						'where we make a quick preview of the data generated in form
							thisfield=getfield(getums, t)
							thisfieldtype=getfieldtype(getums, t)
							if (thisfieldtype="!mg" or thisfieldtype="!mr" or thisfieldtype="!ml") and wehavedata then
								imagearray=split(thisfield, "|")
								if ubound(imagearray)=2 then
									thisfield=imagearray(0)
									imagewidthout=imagearray(1)
									imageheightout=imagearray(2)
								end if
							else
								imagewidthout=imagewidth
								imageheightout=imageheight
							end if
							'Response.Write(imagewidth & " " & imageheight)
							if thisfield<>"" then
								select case thisfieldtype
									case "img"
										response.write("<img src="& thisfield & " width="& imagewidthout &" height="& imageheightout &" >")
									case "tbx"
										response.write("<font face=arial><b>" & thisfield & "</b></font><br>")
									case "tit"
										response.write("<center><h2><font face=arial>" & thisfield &"</font></h2></center>")
									case "shd"
										response.write("<h2><font face=arial>" & thisfield &"</font></h2>") 
									case "lhd"
										response.write("<font face=arial><b>" & thisfield & "</b></font><br>")
									case "sig"
										response.write("<font face=arial><b>--" & thisfield & "</b></font><br>")
									case "snd"
										response.write("<embed src="& thisfield & " width=100 height=100 ><br>")
									case "ltx"
										response.write("<table bgcolor=" & medium & " align=left width=100><font face=arial>" & thisfield & "</font></td></table>")
									case "rtx"
											response.write("<table bgcolor=" & medium & " align=right width=100><font face=arial>" & thisfield & "</font></td></table>")
									case "iml"
										response.write("<img src="& thisfield & " width=" & imagewidthout & " height=" & imageheightout & " align=left>")
									case "imr"
										response.write("<img src="& thisfield & " width=" & imagewidthout & " height=" & imageheightout & " align=right>")
									case "qut"
										response.write("<center><h2><font face=arial>" & thisfield &"</font></h2></center>")
									case "qus"
										response.write("<font face=arial>" & thisfield & "</font><br>")
									case "qul"
										response.write("<font face=arial>" & thisfield & "</font><br>")
									case "ans"
										response.write("<p><b><font face=arial>" & thisfield & "</font></b><br>")
									case else
										response.write("")
			   
								end select
							end if
					next
		
		%>	
			
			<p align=right>
			
<form action=front.asp>
<input type=hidden name=tool value=<%=tool%>>
<select name=FileDatePart><%
for xx=0 to intFileTop
	if handlekeys then
		if len(file.name)=20 then
			if keyeddatepart=left(arrFile(xx), 8) then
				okayfornextpart=true
			else
				okayfornextpart=false
			end if
		else
			okayfornextpart=false
		end if
	else
		okayfornextpart=true
	end if
	if isnumeric(left(arrFile(xx), 8)) and okayfornextpart then
		set fileEntry=fso.opentextfile(server.mappath(strTopPath & bottommostdirectory & "\"& keyeddatepart & arrFile(xx) & ".htm"), 1)
		getums=fileEntry.readall
		fileEntry.close
		whatweshow=getfield(getums, cint(fieldforpulldown))
		'parse out the pieces using the new scheme
		lastfieldnumber=getlastfieldnumber(getums)
		if not handlekeys then 
			if arrFile(xx)=FileDatePart then
				sel="selected"
				posselected=posselected+1
			else
				sel=""
			end if
		else
	   		if FileDatePart=FileDatePart then
				sel="selected"
				posselected=posselected+1
			else
				sel=""
			end if
		end if			
		if not handlekeys then
			response.write("<option " & sel & " value=" & arrFile(xx) & ">"& whatweshow &"<br>")
		else
			response.write("<option " & sel &" value=" & arrFile(xx) & ">"& whatweshow &"<br>")
		end if
	end if

next
%></select>
<input type=submit name=xx value="Let's Go!">
</form>
			
			
			<font  size=1 face=arial,helvetica>
			<%
			if placement-1=>0 then %>
				<a href=sexrated.asp?keyplace=<%=keyplace%>&placement=<%=placement-1%>&status=<%=status%>>earlier responses</a>
			<%end if
			
			if placement+1<=intFileTop and placement-1=>0 then
				if  makedate(arrFile(placement+1))<=date then
					response.write(" | ")
				end if 
			end if
			if placement+1<=intFileTop then
				'response.write(makedate(arrFile(placement+1))&" "&date)
				if makedate(arrFile(placement+1))<=date then
					%>
					<a href=sexrated.asp?keyplace=<%=keyplace%>&placement=<%=placement+1%>&status=<%=status%>>later responses</a>
				<%end if
			end if
			%>
			<p>
			</font></td>
			</tr>
			<%if status<>"thankyou" then%>
				<tr bgcolor="<%=medium%>"> 
				<td colspan=2>
				<font size="4" face="Arial, Helvetica, sans-serif" color="<%=bgcolor%>">&nbsp;&nbsp;&nbsp;<%
				
			if keyedto<>"" then%>RESPOND to a <%=keyname%><%else%>FEEDBACK Section<%end if%> </font> </tr>
				<tr><form action=mailer.asp>
				<td colspan=2  bgcolor=<%=white%>><div align=center><font face=arial,helvetica size=2>
				<br><%=responseintro%><p><font size=3>
				<%=question%></font><font size=2>
				<p>
				<input type=hidden name=responseemail value="<%=responseemail%>">
				<input type=hidden name=tool value="<%=tool%>">
				<input type=hidden name=placement value="<%=placement%>">
				<input type=hidden name=keyplace value="<%=keyplace%>">
				<input type=hidden name=question value="<%=question%>">
				<input type=hidden name=shortquestion value="<%=headersmall%>">
				<textarea name=answer wrap=physical ROWS="18" COLS="32"></textarea><br>
				signed:<input type=text name=signed value="" size=12>
				<font size=1>
				<input type=submit name=vava value="Respond">
				</font><p>&nbsp;
				</div>
				</td>
				</tr>
			<%else%>
				<tr bgcolor="<%=light%>"> 
				<td colspan=2>
				<font size="3" face="Arial, Helvetica, sans-serif" color="<%=medium%>"><b>&nbsp;&nbsp;&nbsp;By the way, thanks for answering our question! </font> </tr>
				<tr><form action=mailer.asp method=post name="GusWorksHere">
				<td colspan=2  bgcolor=<%=white%>> &nbsp;<blockquote><font face=verdana, helvetica size=2>
				Check back here regularly to see if we publish your response!
				<p>&nbsp;<p>
				</font></td>
				</tr>
			<%
			end if
		end if
	end if
end if

%>

<%
if keyedto<>"" then
%>
	<tr bgcolor="<%=light%>"> 
	<td colspan=2>
	<font size="4" face="Arial, Helvetica, sans-serif" color="<%=medium%>">&nbsp;&nbsp;&nbsp;<%if status<>"new" then%>OTHER<%else%>Index of<%end if%> <b>SEX RATED</b> Questions </font></tr>
	<tr>
	<td colspan=2 bgcolor=white><font face=arial,helvetica size=2><p>&nbsp;<p>
	<%
	redim index(qatop, 1) 
	toppest=0
	topmost=0
	for t=qatop to 0 step-1
		set questionfullfile=fso.opentextfile(server.mappath(strKeyedPath & "\" &arrKey(t)&".htm"), 1, true)
		getums=trim(questionfullfile.readall)
		getums=replace(getums, chr(13), "")
		getums=replace(getums, chr(10), "")
		questionfullfile.close
		if instr(getums, "|")<>0 then
			stor=split(getums, "|")
			headersm=stor(0)
			questionation=stor(1)
		else 
			questionation=getums
		end if
		if instr(questionation, "\")<>0 then
			stor=split(questionation, "\")
			questionation=stor(0)
			image=stor(1)
		else
			image=""
		end if
		if status="new" or t<>cint(keyplace) then
			if status="new" then
				'calculate which question has the latest answer
				set answercat=fso.opentextfile(server.mappath(strTopPath &bottommostdirectory & "/" & arrKey(t) & "catalog.htm"), 1)
				getums=trim(answercat.readall)
				answercat.close
				if right(getums, 1)="|" then
					getums=left(getums, len(getums)-1)
				end if
				catal=split(getums, "|")
				topofcatalog=catal(ubound(catal))
				if topofcatalog>topmost then
					topmost=topofcatalog
					toppest=t
				end if
				if image<>"" then
					index(t,0)="<a href=sexrated.asp?keyplace="&t&"><img src="&image&" width=72 height=72 border=0 align=left></a>"
				end if
			end if
			index(t,0)=index(t,0)+"&nbsp;<a href=sexrated.asp?keyplace="&t&"><b>"& headersm &"</b></a><br><dd>"& questionation &"<br clear=all>"
			if (status<>"new") or (t<=qatop and status="new" and not t=0) then
				index(t,1)=index(t,1)+"<p><hr noshade size=1 width=33% ><p>"
				linecount=linecount+1
			end if
		end if
	next
	'output the stuff
	for t=qatop to 0 step-1
		if t=toppest and status="new" then
			response.write("<table align=right bgcolor=" & grey & "><td><font color=white face=arial><b>Updated Today!</td></table>")
		end if
		response.write(index(t,0))
		if index(t,1)<>"" then
			possibleline=possibleline+1
		end if
		if possibleline<linecount or status="new" then
			response.write(index(t,1))
		end if
	next 
	%>
	<p>&nbsp;<p>
	</td><tr>
<%end if 

%>
</table></td>
 </tr>

</table>

&nbsp;<p>


<P>
</BODY>
</HTML>
<%

function makedate(datron)
fakeyear=left(datron, 4)
fakemonth=mid(datron, 5, 2)
fakeday=mid(datron, 7, 2)
if isdate(fakemonth&"/"&fakeday&"/"&fakeyear) then
makedate=cdate(fakemonth&"/"&fakeday&"/"&fakeyear)
else 
makedate=""
end if
end function

function readlines(number)
'reads time remaining for a specific action using the FSO
	set timefile=fso.opentextfile(server.mappath(toolpath & "/" & tool & ".tol"), 1, true)
	if number => 1 then
		for t=0 to number-1
			if not timefile.atendofstream then
				getums=timefile.readline
			else 
				getums="$devoid$"
			end if
		next
	end if
	if not timefile.atendofstream then
		getums=timefile.readline
	else
		getums="$devoid$"
	end if
	timefile.close
	readlines=trim(getums)
end function

function cleanup(inner)
'gets rid of junk chr(13)s, etc.
	if inner<>"" then
	inner=replace(inner, chr(10), "")
	cleanup=trim(replace(inner, chr(13)&chr(10)&chr(13)&chr(10), chr(13)&chr(10)))
		do while instr(cleanup, "  ") or instr(cleanup, chr(13) & chr(13)) 
			cleanup=replace(cleanup, chr(13)&chr(13), chr(13))
			cleanup=replace(cleanup, "  ", " ")
		loop
		if left(cleanup, 1)=chr(13) then cleanup=trim(right(cleanup, len(cleanup)-1))
		if cleanup<>"" then
			if right(cleanup, 1)=chr(13) then cleanup=trim(left(cleanup, len(cleanup)-1))
		end if
	else
		cleanup=""
	end if
end function

function getfield(inputstring, itemnumber)
	if inputstring<>"" then
		start=1
		do while not done and start>0
			prelimer=instr(start, inputstring, "<$")
			if prelimer<>0 then
				localfieldnumber=cint(mid(inputstring, 5+prelimer, 2))
				if localfieldnumber = itemnumber then
					fieldstart=prelimer+8
					nextbracket=instr(fieldstart, inputstring, "<$")
					if nextbracket=0 then
						fieldend=len(inputstring)+1
					else
						fieldend=nextbracket
					end if
					getfield=mid(inputstring, fieldstart, fieldend-fieldstart)
					done=true
				end if
				start=instr(prelimer+6, inputstring, "<$")
			end if
		loop
		if not done then
			getfield="empty"
		end if
	else
		getfield=""
	end if
end function

function getfieldtype(inputstring, itemnumber)
	start=1
	inputstring=cleanup(inputstring)
	if inputstring<>"" then
		do while not done and start
			prelimer=instr(start, inputstring, "<$")
			if prelimer<>0 then
				localfieldnumber=cint(mid(inputstring, 5+prelimer, 2))
				if localfieldnumber = itemnumber then
					typestart=prelimer+2
					getfieldtype=mid(inputstring, typestart, 3)
					done=true
				end if
				start=instr(prelimer+6, inputstring, "<$")
			end if
		loop
		if not done then
			getfieldtype="xxx"
		end if
	else
		getfieldtype="yyy"
	end if
end function

function getlastfieldnumber(inputstring)
	if len(inputstring)>7 then
		numberpos=instrrev(inputstring, "<$")+5
		if numberpos=0 then 
			Getlastfieldnumber=-1
		else
			strGetlastfieldnumber=mid(inputstring, numberpos, 2)
			if isnumeric(strGetlastfieldnumber) then
				Getlastfieldnumber=cint(strGetlastfieldnumber)
			else
				Getlastfieldnumber=0
			end if
		end if
	else
		Getlastfieldnumber=-1
	end if
end function

function backslash(inner)
	backslash= replace(inner, "/" , "\")
end function

function forwardslash(inner)
	forwardslash= replace(inner, "\" , "/")
end function

function filetoplacement(strFilein)
for xx=0 to intFileTop
	if cstr(strFilein)=cstr(arrFile(xx)) then
		filetoplacement=xx
	end if
next
if filetoplacement="" then
	filetoplacement=0
end if
end function
%>


<!--#include virtual="/include/CCFooter.asp"-->
