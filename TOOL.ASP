<!--#include virtual="/include/CCHeader.asp"-->

<%
'another Gus Mueller automaton
'February 26 1999
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
ho=cstr(hour(time))
if len(ho)=1 then 
	ho = "0" & ho
end if
mi=cstr(minute(time))
if len(mi)=1 then 
	ho = "0" & ho
end if
set fso=server.createobject("scripting.filesystemobject")
i=0
tool=request("tool")
if tool="" then response.redirect("toolbuilder.asp")
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
do until readlines(i)="$devoid$" or readlines(i)=""
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

strTopPath=topmostdirectory & "\" & strShortPath
strKeyedPath=topmostdirectory & "\" & keyedto

itemvalid=0
redim formitems(fieldnumber, 2)
for i=0 to fieldnumber
	formitems(i, 0)=request("field"& i)
	formitems(i, 1)=fieldtype(i)
	if formitems(i,0)<>"" then
		itemvalid=itemvalid+1
	end if
next 

'make file structures as needed
if not fso.folderexists(server.mappath(topmostdirectory)) then
	fso.createfolder server.mappath(topmostdirectory)
end if
if not fso.folderexists(server.mappath(topmostdirectory & "\" & strShortPath)) then
	fso.createfolder server.mappath(topmostdirectory & "\" & strShortPath)
end if
if not fso.folderexists(server.mappath(topmostdirectory & "\" & strShortPath & bottommostdirectory)) then
	fso.createfolder server.mappath(topmostdirectory & "\" & strShortPath  & bottommostdirectory)
end if

%>
<html>
<head>
<title><%=strToolName%></title>
</head>
<body bgcolor=<%=bgcolor%> text=<%=text%>>
<font face=arial><font face=arial>
<form action=tool.asp>
Other Tools:
	<select name=tool>
	<option value="">Tool Builder
	<%
set fold=fso.getfolder(server.mappath(toolpath))
set filies=fold.files
for each fy in filies
keyer=left(fy.name, len(fy.name)-4)
if right(fy.name, 3)="tol" then
	if tool=keyer then 
		sel="selected"
	else
		sel="" 
	end if
	response.write("<option "&sel&" value=" & keyer & ">" & keyer )
end if
next
	%>
	</select><Input type=submit value="Run"></form>
<h2><%=strToolName%>: <% if FileDatePart<>"" then response.write(left(FileDatePart,4)&"-"&mid(FileDatePart,5,2)&"-"&right(FileDatePart,2))%></h2>
Use this form to edit <%=strToolName%> entries.
<p>
<%
if changedate="" then 
	moder="write"
else
	moder="changedate"
end if 
if FileDatePart<>"" then 
	if itemvalid>0 then
	'response.write(filebackupname & " " & filedatepart)
		if moder="write" and fileBackupName=FileDatePart and deletion="" then
		'response.write(strTopPath & bottommostdirectory & "\"&FileDatePart&".htm")
			stringtosave=""
			for i=0 to fieldnumber
				imagewidth=request("width" & i)
				imageheight=request("height" & i)
			if formitems(i, 1)="!mg" or formitems(i, 1)="!mr" or formitems(i, 1)="!ml" then
				fieldtoadd=formitems(i,0) & "|" & imagewidth & "|" & imageheight
			else
				fieldtoadd=formitems(i,0)
			end if
				stringtosave=addfield(stringtosave, fieldtoadd, formitems(i, 1))
			next 
			set fileEntry=fso.createtextfile(server.mappath(strTopPath & bottommostdirectory & "\"&keyeddatepart&FileDatePart&".htm"), true)
			fileEntry.write(stringtosave)
			fileEntry.close

			response.write("<table bgcolor=ffffff width=400><td><font face=arial size=2 color="& text &">")
			for t=0 to fieldnumber
			'where we make a quick preview of the data generated in form
				thisfield=formitems(t, 0)
				thisfieldtype=formitems(t, 1)
				if thisfieldtype="!mg" or thisfieldtype="!mr" or thisfieldtype="!ml" then
				thisfieldtype="i" & right(thisfieldtype, 2)
				imagewidth=request("width" & t)
				imageheight=request("height" & t)
				'Response.Write(imagewidth & " " & imageheight)
				end if
				if thisfield<>"" then
					select case thisfieldtype
						case "img"
							response.write("<img src="& thisfield & " width="& imagewidth &" height="& imageheight&" >")
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
							response.write("<img src="& thisfield & " width="&imagewidth&" height="&imageheight&" align=left>")
						case "imr"
							response.write("<img src="& thisfield & " width="&imagewidth&" height="&imageheight&" align=right>")
						case "qut"
							response.write("<center><h2><font face=arial>" & thisfield &"</font></h2></center>")
						case "qus"
							response.write("<font face=arial>" & thisfield & "</font><br>")
						case "qul"
							response.write("<font face=arial>" & thisfield & "</font><br>")
						case "ans"
							response.write("<p><b><font face=arial>" & thisfield & "</font></b><br>")
						case else
							response.write("<b>WHOOPS!!</b>")

					end select
				end if
			next
			response.write("</td></table><br>was written to Dev.</font>")
			if not fso.fileexists(server.mappath("\trans\"&mo&da & strShortPath & ".tra")) then
				set transcript=fso.createtextfile(server.mappath("\trans\"&mo&da& strShortPath & ".tra"), true)
				transcript.close
			end if
			
			set transcript=fso.opentextfile(server.mappath("\trans\"&mo&da &  strShortPath & ".tra"), 8)
			transcript.writeline("\ccroot" & backslash(strTopPath) & bottommostdirectory & "\" & keyeddatepart&FileDatePart&".htm")
			transcript.close
			response.write("<form method=post action=http://content/update/copy.asp><font size=+1 color=red>" & strShortPath &" has been modified</font><br>")
			response.write(strShortPath & " file:<input type=text name=schedfile value="& mo & da & strShortPath & ".tra><INPUT TYPE=SUBMIT value="&chr(34)&"Copy the files to the live server"&chr(34)&"><input type=hidden name=username value="&chr(34)& strShortPath & " front end"&chr(34)&"></form>")
		
		end if 'moder=write
	end if
end if 'FileDatePart

%>
<form method="post" action="tool.asp" name=beer>
<input type=hidden name=tool value=<%=tool%>>
<%
if handlekeys then
	if keyeddatepart="" and FileDatePart<>"" then
		FileDatePart=""
	end if
	'time to generate a menu of keying options for keyed pods
	if fso.folderexists(server.mappath(strKeyedPath & bottommostdirectory)) then
		%>
		<b><%=keyname%>:</b><font size=1>
		<select name=<%=keyeddatepart%>>
			
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
	<a href=faqmaker.asp?tool=<%=fak%>&changedate=EE&fileBackupName=<%=FileDatePart%>&FileDatePart=<%=FileDatePart%>>edit</a>
	</font><br>
	<%
	else
		response.write("Keyed Tool not built! ")
		response.write("<b><a href=toolbuilder.asp?tool="& keyedto &"> Build " & keyname & "</a>?</b>")
		response.end
	end if
	%>
	<table bgcolor=white><td><font face=arial size=3><b><%=bigwhatweshow%></b></td></table>
	
	
	<%
end if
if filedatepart="" then
	FileDatePart=ye&mo&da
end if


if deletion<>"" then
response.write("<font color=990000>Deleted: " &strTopPath & bottommostdirectory & "\"& keyeddatepart& FileDatePart & ".htm</font>")
	if fso.fileexists(server.mappath(strTopPath & bottommostdirectory & "\"& keyeddatepart& FileDatePart & ".htm")) then
		fso.deletefile(server.mappath(strTopPath & bottommostdirectory & "\" & keyeddatepart & FileDatePart & ".htm"))
	end if
end if

if moder="changedate" then
	if fso.fileexists(server.mappath(strTopPath & bottommostdirectory & "\" & keyeddatepart & FileDatePart&".htm")) then
		set fileEntry=fso.opentextfile(server.mappath(strTopPath & bottommostdirectory & "\" & keyeddatepart &FileDatePart&".htm"), 1)
		getums=fileEntry.readall
		fileEntry.close
		getums=cleanup(getums)
		'parse out the pieces using the new scheme
		lastfieldnumber=getlastfieldnumber(getums)
		wehavedata=true
	else 
		response.write("<br><font color=red>note: this entry didn't exist.</font><br>")
		wehavedata=false
		lastfieldnumber=fieldnumber
		getums=""
	end if
	redim widthe(lastfieldnumber)
	redim heighte(lastfieldnumber)
	response.write("<div align=right><font size=1><input name=delete type=checkbox> Delete this entry</font></div><br>")
	for t = 0 to lastfieldnumber
	'where we generate the form based on pod configuration data
		if wehavedata then
			thisfieldtype=getfieldtype(getums, t)
			thisfield=getfield(getums, t)
		else
			thisfieldtype=fieldtype(t)
			thisfield=""
		end if
				if (thisfieldtype="!mg" or thisfieldtype="!mr" or thisfieldtype="!ml") and wehavedata then
					imagearray=split(thisfield, "|")
					if ubound(imagearray)=2 then
						thisfield=imagearray(0)
						widthe(t)=imagearray(1)
						heighte(t)=imagearray(2)
					end if
				end if
			select case thisfieldtype
				case "img"
					response.write("<b>image path:</b>")
					response.write("<input name=field"& t &" size=16 value="& chr(34) & thisfield & chr(34) &"><br>")
				case "tbx"
					response.write("<textarea name=field"& t &" ROWS=6 COLS=55 wrap=soft >" &thisfield & "</textarea><br>")
				case "tit"
					response.write("<b>title:</b>")
					response.write("<input name=field"& t &" size=16 value=" & chr(34) & thisfield & chr(34) &"><br>")
				case "shd"
					response.write("<b>short header:</b>")
					response.write("<input name=field"& t &" size=16 value=" &  chr(34) & thisfield & chr(34) &"><br>")
				case "lhd"
					response.write("<b>long header:</b>")
					response.write("<textarea name=field"& t &" ROWS=2 COLS=44 wrap=soft >" &thisfield & "</textarea><br>")
				case "sig"
					response.write("<b>signed:</b>") 
					response.write("<input name=field"& t &" size=16 value="  & chr(34) & thisfield & chr(34) &"><br>")
				case "snd"
					response.write("<b>sound path:</b>") 
					response.write("<input name=field"& t &" size=16 value=" &  chr(34) & thisfield & chr(34) &"><br>")
				case "ltx"
					response.write("<b>left-aligned text box:</b><br>")
					response.write("<textarea name=field"& t &" ROWS=6 COLS=55 wrap=soft >" &thisfield & "</textarea><br>")
				case "rtx"
					response.write("<b>right-aligned text box:</b><br>")
					response.write("<textarea name=field"& t &" ROWS=6 COLS=55 wrap=soft >" &thisfield & "</textarea><br>")
				case "iml"
					response.write("<b>left-aligned image path:</b>")
					response.write("<input name=field"& t &" size=16 value=" &  chr(34) & thisfield & chr(34) &"><br>")
				case "imr"
					response.write("<b>right-aligned image path:</b>")
					response.write("<input name=field"& t &" size=16 value=" &  chr(34) & thisfield & chr(34) &"><br>")
				case "qut"
					response.write("<b>Question (for menu):</b>")
					response.write("<input name=field"& t &" size=16 value=" & chr(34) & thisfield & chr(34) &"><br>")
				case "qus"
					response.write("<b>Question:</b>")
					response.write("<textarea name=field"& t &" ROWS=2 COLS=44 wrap=soft >" &thisfield & "</textarea><br>")
				case "qul"
					response.write("<b>Question:</b><br>")
					response.write("<textarea name=field"& t &" ROWS=8 COLS=55 wrap=soft >" &thisfield & "</textarea><br>")
				case "ans"
					response.write("<b>Answer:</b><br>")
					response.write("<textarea name=field"& t &" ROWS=8 COLS=55 wrap=soft >" &thisfield & "</textarea><br>")
				case "!mr"
					response.write("<b>right-aligned image path:</b>")
					response.write("<input name=field"& t &" size=16 value=" &  chr(34) & thisfield & chr(34) &"><br>")
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;width: <input name=width"& t &" size=4 value=" &  chr(34) & widthe(t) & chr(34) &">")
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;height: <input name=height"& t &" size=4 value=" &  chr(34) & heighte(t) & chr(34) &"><br>")
				case "!ml"
					response.write("<b>left-aligned image path:</b>")
					response.write("<input name=field"& t &" size=16 value=" &  chr(34) & thisfield & chr(34) &"><br>")
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;width: <input name=width"& t &" size=4 value=" &  chr(34) & widthe(t) & chr(34) &">")
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;height: <input name=height"& t &" size=4 value=" &  chr(34) & heighte(t) & chr(34) &"><br>")
				case "!mg"
					response.write("<b>image path:</b>")
					response.write("<input name=field"& t &" size=16 value=" &  chr(34) & thisfield & chr(34) &"><br>")
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;width: <input name=width"& t &" size=4 value=" &  chr(34) & widthe(t) & chr(34) &">")
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;height: <input name=height"& t &" size=4 value=" &  chr(34) & heighte(t) & chr(34) &"><br>")
				case else
					response.write("<b>WHOOPS!!:</b>")
					response.write("<input name=field"& t &" size=16 value=" & chr(34) & thisfield & chr(34) &"><br>")
			end select
	next
end if
if (handlekeys and keyeddatepart<>"") or (not handlekeys) then
	%>
	<p><select name=FileDatePart onChange="document.beer.changedate.click();">
	<%
	'response.write(server.mappath(strTopPath & bottommostdirectory))
	posselected=0
	if isnumeric(fieldforpulldown) then
		buttondescription="Select Entry"
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
				%>
					<option <%=sel%> value=<%=left(file.name, len(file.name)-4)%>><%=whatweshow%>
				<%
				else
				%>
					<option <%=sel%> value=<%=left(file.name, mid(file.name, 9, 8))%>><%=whatweshow%>
				<%
				end if
			end if
		next
		if posselected=0 then
			response.write("<option selected value="& FileDatePart &">"&  "*new entry*")
		end if
		do while fso.fileexists(server.mappath(strTopPath & bottommostdirectory & "\"& keyeddatepart & ye&mo&da&".htm"))  
			newdate=dateadd("d", 1, makedate(ye&mo&da))
			da=day(newdate)
			mo=month(newdate)
			ye=year(newdate)
			if len(mo)=1 then 
				mo = "0" & mo
			end if
			if len(da)=1 then 
				da = "0" & da
			end if
		loop
		if ye&mo&da=FileDatePart then 
			sel="selected"
		else
			sel=""
		end if
		if posselected<>0 then
			response.write("<option " & sel & " value=" & ye&mo&da & ">*new entry*")
		end if
		response.write("</select> entry&nbsp;")
	else
		buttondescription="Select Day"
		olddate=makedate(FileDatePart)
		if datediff("d", olddate, date) > 7 then
			%>
			<option  value=<%=FileDatePart%>><%=left(weekdayname(weekday(olddate)),2)%>  <%=month(olddate)%>/<%=day(olddate)%>/<%=year(olddate)%>*
			<%
		end if 
		
		for daze=-7 to 21
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
			if ye&mo&da=FileDatePart then
				sel="selected"
			else
				sel=""
			end if
			'look to see if there is a legitimate file here
			ex=""
			if fso.fileexists(server.mappath(strTopPath & bottommostdirectory & "\"& keyeddatepart & ye&mo&da&".htm")) then
				set ourfile=fso.getfile(server.mappath(strTopPath & bottommostdirectory & "\"& keyeddatepart &ye&mo&da&".htm"))
					ourfilesize=ourfile.size
				if ourfilesize>4 then 
					ex="*"
				end if
			end if
			
			%>
			<option <%=sel%> value=<%=ye&mo&da%>><%=left(weekdayname(weekday(datem)), 2)& " "%><%=mo&"/"&da&"/"&right(ye,2)&ex%>
		<%next%>
		</select> day&nbsp;
<%end if
end if

if changedate="" and writer="" then
buttondescription="Select " & keyname
end if
%>
<INPUT TYPE=SUBMIT value="<%=buttondescription%>" name="changedate">
<%if moder<>"write" then%>
<INPUT TYPE=SUBMIT value="Write to Dev" name="writer">
<%end  if%>
<INPUT TYPE=hidden value="<%=FileDatePart%>" name="fileBackupName">
</form>
<p>
<ul>
<font face=arial>
<p>
When you're done making a batch of <%=strToolName%> articles, you should<br>
<b> <a href=cataloger.asp?keyeddatepart=<%=keyeddatepart%>&FileDatePart=<%=FileDatePart%>&tool=<%=tool%>>Update</a> the <%=strToolName%> Past Entries.</b><p>
Go check out the <a href=<%=forwardslash(frontendurl)%>><%=strToolName%> front end</a>.
<p>
 <%=strToolName%> entry form</a> | 
 <a href=cataloger.asp?keyeddatepart=<%=keyeddatepart%>&FileDatePart=<%=FileDatePart%>&tool=<%=tool%>><%=strToolName%> catalog updater</a>
</P>
</body>
</html>

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
	'if not fso.folderexists(server.mappath(toolpath & "/" & tool & ".tol")) then response.redirect("toolbuilder.asp")
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
