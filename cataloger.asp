<!--#include virtual="/include/CCHeader.asp"-->

<%
'another Gus Mueller automaton
'Jan 1999
set fso=server.createobject("scripting.filesystemobject")
tool=request("tool")
if tool="" then tool ="faq"
defaulttopmost="\extension"
toolfolder="tools"
toolpath=defaulttopmost & "\" & toolfolder

'read all data specific to this implementation<br>
bgcolor=readlines(0)
medium=readlines(1)
text=readlines(2)
link=readlines(3)
vlink=readlines(4)
tool=readlines(5)
strToolName=readlines(6)
strShortPath=readlines(7)
fieldnumber=readlines(8)
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
i=30

if fieldforpulldown="" then	 fieldforpulldown=0
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

deletion=request("delete")
changedate=request("changedate")
fileBackupName=request("fileBackupName")
writer=request("writer")
FileDatePart=request("FileDatePart")
keyeddatepart=request("keyeddatepart")
if keyedto<>"" then
	handlekeys=true
else
	handlekeys=false
end if
%>
<html><head>
<title><%=strToolName%> Catalog updater</title>
</head><body bgcolor=<%=bgcolor%> text=<%=text%>><font face=arial size=3>
<font face=arial>
<font size=-1>
</font>
<p>
<ul>
<h2>updating <%=strToolName%> Catalog...</h2>
<b>files (click to edit):</b><br>
<%
strTopPath=topmostdirectory & "\" & strShortPath
strKeyedPath=topmostdirectory & "\" & keyedto

fresh=1
ye=cstr(year(date))
mo=cstr(month(date))
if len(mo)=1 then 
mo = "0" & mo
end if
da=cstr(day(date))
if len(da)=1 then 
da = "0" & da
end if
FileDatePart=request("FileDatePart")
'response.write(strTopPath & bottommostdirectory & keyeddatepart & "\" & "catalog.htm")
set catalog=fso.createtextfile(server.mappath(strTopPath & bottommostdirectory & keyeddatepart & "\"  & "catalog.htm"), true)
set filez=fso.getfolder(server.mappath(strTopPath & bottommostdirectory)).files
for each file in filez
	if handlekeys then
		if len(file.name)=20 then
			if keyeddatepart=left(file.name, 8) then
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
	if isnumeric(left(file.name, 8)) and okayfornextpart then

		set fileEntry=fso.opentextfile(file.path, 1)
		getums=fileEntry.readall
		fileEntry.close
		whatweshow=getfield(getums, cint(fieldforpulldown))
		'parse out the pieces using the new scheme
		lastfieldnumber=getlastfieldnumber(getums)
		if not handlekeys then 
			if left(file.name, len(file.name)-4)=FileDatePart then
				sel="selected"
				posselected=posselected+1
			else
				sel=""
			end if
		else
	   		if mid(file.name, 9,  8)=FileDatePart then
				sel="selected"
				posselected=posselected+1
			else
				sel=""
			end if
		end if			
		if not handlekeys then
			response.write("<a href=tool.asp?tool="& tool &"&changedate=Select&keyeddatepart="& keyeddatepart &"&fileDatePart=" & left(file.name, len(file.name)-4)  & "&fileBackupName=" & left(file.name, len(file.name)-4) &">" &  file.name & "</a> "& whatweshow &"<br>")
			catalog.write(left(file.name, len(file.name)-4)& "|" )
		else
			response.write("<a href=tool.asp?tool="& tool &"&changedate=Select&fileDatePart=" & left(file.name, len(file.name)-4)  & "&fileBackupName=" & left(file.name, len(file.name)-4) &">" &  file.name & "</a> "& whatweshow &"<br>")
catalog.write(left(file.name, len(file.name)-4)& "|" )
		end if
	end if
next
catalog.close
%>
<p>

<form action=cataloger.asp>
<input type=hidden name=tool value=<%=tool%>>
<% 
if handlekeys then
%>
<b><%=keyname%>:</b><font size=1>
	<select name=keyeddatepart>
<%
	buttondescription="Select Entry"
	set filez=fso.getfolder(server.mappath(strKeyedPath & bottommostdirectory)).files
	for each file in filez
		if isnumeric(left(file.name, 8)) then
			set fileEntry=fso.opentextfile(file.path, 1)
			getums=fileEntry.readall
			fileEntry.close
			whatweshow=getfield(getums, cint(fieldforkeyedpulldown))
			if left(file.name, len(file.name)-4)=keyeddatepart then
				sel="selected"
				bigwhatweshow=getfield(getums, cint(fieldforbigkeyedpulldown))
			else
				sel=""
			end if
			%>
			<option <%=sel%> value=<%=left(file.name, len(file.name)-4)%>><%=whatweshow%>
			<%
		end if
	next
	
%>
	</select> entry&nbsp;
	<%
end if
%>

<input type=submit value="Update <%=strShortPath%> catalog">
</form><p></ul>
<font size=-1 face=arial>
</font>
</html>
<%
if not fso.fileexists(server.mappath("\trans\"&mo&da&strShortPath &".tra")) then
set transcript=fso.createtextfile(server.mappath("\trans\"&mo&da & strShortPath &".tra"), true)
transcript.close
end if

set transcript=fso.opentextfile(server.mappath("\trans\"&mo&da & strShortPath &".tra"), 8)
transcript.writeline("\ccroot" & strTopPath & bottommostdirectory & keyeddatepart & "\"  & "catalog.htm")
transcript.close
%>
<p>
Go check out the <a href=<%=forwardslash(frontendurl)%>?FileDatePart=<%=FileDatePart%>&answerfile=<%=answerfile%>><%=strShortPath%> front end</a>.
<p>
 <a href=tool.asp?keyeddatepart=<%=keyeddatepart%>&FileDatePart=<%=FileDatePart%>&tool=<%=tool%>><%=strToolName%> entry form</a> | 
 <%=strToolName%> catalog updater
</P>


<%
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

function backslash(inner)
	backslash= replace(inner, "/" , "\")
end function

function forwardslash(inner)
	forwardslash= replace(inner, "\" , "/")
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

function addfield(inputstring, field, fieldtype)
	if fieldtype="" then
		fieldtype="txl"
	end if
	lastfieldnum=getlastfieldnumber(inputstring)
	if lastfieldnum=-1 then
		addfield="<$" & left(fieldtype, 3) & "00" & ">" & field
	else
		thisfieldnum=lastfieldnum+1
		if len(thisfieldnum)=1 then 
			thisfieldnum="0" & thisfieldnum
		end if
		addfield=inputstring & "<$" & left(fieldtype, 3) & thisfieldnum & ">" & field
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
%>

<!--#include virtual="/include/CCFooter.asp"-->
