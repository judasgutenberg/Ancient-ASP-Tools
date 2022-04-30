<!--#include virtual="/include/CCHeader.asp"-->

<%
'Gus Mueller, March 1999
'this robot builds the configurations for 
'all of the "pod builder" robots

defaulttopmost="\extension"
defaultbottommost="\entries"
toolfolder="tools"
toolpath=defaulttopmost & "\" & toolfolder
toolgo=request("toolgo")
select case request("go")
	case "Run"
		response.redirect("tool.asp?tool=" & toolgo)
	case "Edit"  
		mode=edit
		tool=toolgo
	case "Build Tool" 
		tool=request("tool")
end select

set fso= server.createobject("scripting.filesystemobject")
'make file structures as needed
if not fso.folderexists(server.mappath(defaulttopmost)) then
	fso.createfolder server.mappath(defaulttopmost)
end if
if not fso.folderexists(server.mappath(toolpath)) then
	fso.createfolder server.mappath(toolpath)
end if

numberoffields=7

if request("go")="Build Tool" and tool<>"" and tool<>"$devoid$" and mode<>"edit" then
	bwlWriter=true
	set toolconfig=fso.createtextfile(server.mappath("/extension/tools/" & tool & ".tol"), true)
	toolconfig.writeline(request("bgcolor"))
	toolconfig.writeline(request("medium"))
	toolconfig.writeline(request("text"))
	toolconfig.writeline(request("link"))
	toolconfig.writeline(request("vlink"))
	toolconfig.writeline(request("tool"))
	toolconfig.writeline(request("strToolName"))
	toolconfig.writeline(tool)
	toolconfig.writeline(request("responseintro"))
	toolconfig.writeline(request("imageheight"))
	toolconfig.writeline(request("imagewidth"))
	toolconfig.writeline(request("frontendurl"))
	toolconfig.writeline(request("keyedto"))
	toolconfig.writeline(request("fieldforkeyedpulldown"))
	toolconfig.writeline(request("fieldforbigkeyedpulldown"))
	toolconfig.writeline(request("keyname"))
	toolconfig.writeline(request("topmostdirectory"))
	toolconfig.writeline(request("bottommostdirectory"))
	toolconfig.writeline(request("responseemail"))
	toolconfig.writeline(request("fieldforpulldown"))
	toolconfig.writeline(request("ftitle"))
	toolconfig.writeline(cleanup(replace(request("introtext"), chr(13), " ")))
	toolconfig.writeline(cleanup(replace(request("thankyou"), chr(13), " ")))
	toolconfig.writeline(cstr(request("communityid")))
	toolconfig.writeline(cstr(request("datefed")))
	toolconfig.writeline(cstr(request("lister")))
	toolconfig.writeline(cleanup(replace(request("brought"), chr(13), " ")))
	toolconfig.writeline(cstr(request("style")))
	toolconfig.writeline(cstr(request("sectionid")))
	'create room for future expansion of this model	
	for xx=0 to 0
		toolconfig.writeline("")
	next
	for xx=0 to numberoffields
		toolconfig.writeline(request("typer" & xx))
	next
		toolconfig.close
end if
	
if fso.fileexists(server.mappath("/extension/tools/" & tool & ".tol")) then
	'read all data specific to this implementation<br>
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
	if isnumeric(datefed) then 
		datefed=cint(datefed)
	else
		datefed=0
	end if
	if isnumeric(lister) then 
		lister=cint(lister)
	else
		lister=0
	end if
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
else
	redim fieldtype(fieldnumber)
	for fields=0 to fieldnumber
		if fields<=ubound(fieldtype) then
		fieldtype(fields)=""
		end if
	next
end if
if frontendurl = defaulttopmost & "\front.asp?tool=" then frontendurl=""
if bgcolor="" then bgcolor="ffffff"
if text="" then text="000000"
if link="" then link="990000"
if vlink="" then vlink="000099"
if medium="" then medium="ccffcc"
if bottommostdirectory = "" then bottommostdirectory=defaultbottommost
if topmostdirectory = "" then topmostdirectory=defaulttopmost
if frontendurl=""  then frontendurl=defaulttopmost & "\front.asp?tool=" & tool
if isnumeric(fieldforpulldown) then
fieldforpulldown=cint(fieldforpulldown)
end if
%>
<html>
<head>
<title>Tool Builder</title>
</head>
<body bgcolor=<%=bgcolor%> text=<%=text%> link=<%=link%> vlink=<%=vlink%>>
<font face=arial>
<h2>Tool Builder 
<%
if strToolName<>"" then
	response.write(": " & strToolName)
end if
%>
</h2>

<form action=toolbuilder.asp method=post>
<table>
<td valign=top>
<table cellspacing=0 cellpadding=3 border=0>
	
	<td bgcolor=<%=medium%>><font face=arial><font color=red>*</font><b>tool short name (no spaces):</td><td bgcolor=<%=medium%>><font face=arial size=2> </b><input type=text size=20 name=tool value="<%=tool%>"></td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial><font color=red>*</font>background color:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2> <input type=text size=20 name=bgcolor value="<%=bgcolor%>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial><font color=red>*</font>medium color:</td><td bgcolor=<%=medium%>><font face=arial size=2><input type=text size=20 name=medium value="<%=medium%>"></td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial><font color=red>*</font>text color:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=text value="<%=text%>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial><font color=red>*</font>link color:</td><td bgcolor=<%=medium%>><font face=arial size=2><input type=text size=20 name=link value="<%=link%>"></td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial><font color=red>*</font>vlink color:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=vlink value="<%=vlink%>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial>tool long name:</td><td bgcolor=<%=medium%>><font face=arial size=2><input type=text size=20 name=strToolName value="<%=strToolName %>"></td><tr>

	<!-- <td bgcolor=<%=bgcolor%>><font face=arial>shortpath:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=strShortPath value="<%=strShortPath%>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial><i>reserved</i>:</td><td bgcolor=<%=medium%>><font face=arial size=2><input type=text size=20 name=reserved value="<%= fieldnumber%>"></td><tr> -->

	<td bgcolor=<%=bgcolor%>><font face=arial>default image height:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=imageheight value="<%= imageheight%>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial>default image width:</td><td bgcolor=<%=medium%>><font face=arial size=2><input type=text size=20 name=imagewidth value="<%=imagewidth %>"></td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial><font color=red>*</font>front end URL:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=frontendurl value="<%= frontendurl%>"></td><tr>

	<td bgcolor=<%=medium%>><font face=arial>keyed to:</td><td bgcolor=<%=medium%>>
	<font face=arial size=2>
	<select name=keyedto>
	<option value="">-nothing-
	<%
set fold=fso.getfolder(server.mappath(toolpath))
set filies=fold.files
for each fy in filies
keyer=left(fy.name, len(fy.name)-4)
if right(fy.name, 3)="tol" and tool<>keyer then
	if keyedto=keyer then 
		sel="selected"
	else
		sel=""
	end if
	response.write("<option "&sel&" value=" & keyer & ">" & keyer )
end if
next
	%>
	</select>
	</td><tr>


	<td bgcolor=<%=bgcolor%>><font face=arial>field number for keyed pulldown:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2>
	<select name=fieldforkeyedpulldown>
	<option value="">date-based
	<%

for xx=0 to numberoffields
if isnumeric(fieldforkeyedpulldown) then fieldforkeyedpulldown=cint(fieldforkeyedpulldown)
if tool<>"" then
	if xx=fieldforkeyedpulldown then 
		sel="selected"
	else
		sel=""
	end if
end if
	response.write("<option "&sel&" value=" & xx & ">" & xx )
next
	%>
	</select>

</td><tr>
	<td bgcolor=<%=medium%>><font face=arial>field for big keyed pullout:</td><td bgcolor=<%=medium%>><font face=arial size=2>


	<select name=fieldforbigkeyedpulldown>
	<option value="">-none-
	<%

for xx=0 to numberoffields
if isnumeric(fieldforbigkeyedpulldown) then fieldforbigkeyedpulldown=cint(fieldforbigkeyedpulldown)
if tool<>"" then
	if xx=fieldforbigkeyedpulldown then 
		sel="selected"
	else
		sel=""
	end if
end if
	response.write("<option "&sel&" value=" & xx & ">" & xx )

next
	%>
	</select>
</td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial>keyname:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=keyname value="<%=keyname %>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial><font color=red>*</font>topmost directory:</td><td bgcolor=<%=medium%>><font face=arial size=2><input type=text size=20 name=topmostdirectory value="<%= topmostdirectory%>"></td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial><font color=red>*</font>bottommost directory:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=bottommostdirectory value="<%=bottommostdirectory %>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial>response email address:</td><td bgcolor=<%=medium%>><font face=arial size=2><input type=text size=20 name=responseemail value="<%= responseemail%>"></td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial><font color=red>*</font>front end title:</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=text size=20 name=ftitle value="<%= ftitle%>"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial>Community ID:<input type=text size=6 name=communityid value="<%= communityid%>"></td><td bgcolor=<%=medium%>><font face=arial>Section ID:<input type=text size=6 name=sectionid value="<%= sectionid%>"></td><tr>
	<td bgcolor=<%=bgcolor%>><font face=arial><input type=checkbox name=datefed <%if datefed=1 then response.write("checked")%> value="1"> date-fed</td><td bgcolor=<%=bgcolor%>><font face=arial size=2><input type=radio name=lister <%if lister=1 then response.write("checked")%> value="1">yes - pulldown index - no<input type=radio name=lister <%if lister=0 then response.write("checked")%> value="0"></td><tr>
	<td bgcolor=<%=medium%>><font face=arial>front end style:</td><td bgcolor=<%=medium%>><font face=arial size=2>
	<select name=style>
	<option <%if style="lovejoy" then response.write("selected")%> value="lovejoy">Dr. LoveJoy
	<option <%if style="styleexperts" then response.write("selected")%> value="styleexperts">Style Experts
	</select>
	</td><tr>
</table>
</td>
<td valign=top>
<table>
<td bgcolor=<%=medium%> colspan=2><font face=arial><font color=red>*</font><b>Field Types:</td><td><font face=arial size=2></b></td><tr>
<td bgcolor=<%=bgcolor%> ><font face=arial size=1><input name=fieldforpulldown type=radio value=""> check this if you want dates instead of field names in your editor's pulldowns</td><td rowspan=11 bgcolor=<%=bgcolor%> width=100><font face=arial size=1>Click the radio buttons beside the field you want to appear in pulldown menus listing these entries.</td><tr>

<%
extra=0
if numberoffields="" then numberoffields=0
if numberoffields > ubound(fieldtype) then 
	extra=numberoffields-ubound(fieldtype)
	numberoffields=ubound(fieldtype)
end if
for t=0 to numberoffields
	formfieldnumber=t
	pulldown
next
for x=1 to extra
	saver=t
	t=0
	saver2=fieldtype(0)
	fieldtype(0)=""
	formfieldnumber=x+t
	pulldown
	fieldtype(0)=saver2
	t=saver
next
%>
</table>
<center>
<input type=submit name=go value="Build Tool">
<%if bwlWriter then%>
	<font face=arial><p>
	<a href=tool.asp?tool=<%=tool%>>Begin Entering Data</a>
<%end if%>

<p><hr noshade size=1><p>

<font face=arial>
<b>Tools:</b>
	<select name=toolgo>
	<option value="">Tool Editor
	<%
set fold=fso.getfolder(server.mappath(toolpath))
set filies=fold.files
for each fy in filies
keyer=left(fy.name, len(fy.name)-4)
if right(fy.name, 3)="tol" then
	response.write("<option "&sel&" value=" & keyer & ">" & keyer )
end if
next
	%>
</select><Input type=submit name=go value="Run"><Input type=submit name=go value="Edit">
	
</td>
</table>
<font face=arial size=2>
<b><font color=red>*</font>front end intro text:</b><br>
<textarea name=introtext ROWS=8 COLS=77 wrap=soft ><%=introtext%></textarea>
<p>
<b>brought to you by:</b><br>
<textarea name=brought ROWS=2 COLS=77 wrap=soft ><%=brought%></textarea>
<p>
<b>response intro:</b><br>
<textarea name=responseintro ROWS=2 COLS=77 wrap=soft ><%=responseintro%></textarea>
<p>
<b>response thank you:</b><br>
<textarea name=thankyou ROWS=2 COLS=77 wrap=soft ><%=thankyou%></textarea>
</form>
</body>
</html>
<%
sub pulldown
%>
	<td><input  <% if formfieldnumber=fieldforpulldown then response.write "checked" %>   name=fieldforpulldown type=radio value=<%=formfieldnumber%>><select name=typer<%= formfieldnumber %>>
	<option value="">-none-
	<option  <% if fieldtype(t)="img" then response.write "selected" %> value="img">image
	<option  <% if fieldtype(t)="imr" then response.write "selected" %> value="imr">image (aligned right)
	<option  <% if fieldtype(t)="iml" then response.write "selected" %> value="iml">image (aligned left)
	<option  <% if fieldtype(t)="!mg" then response.write "selected" %> value="!mg">image (sized)
	<option  <% if fieldtype(t)="!mr" then response.write "selected" %> value="!mr">image (sized, aligned right)
	<option  <% if fieldtype(t)="!ml" then response.write "selected" %> value="!ml">image (sized, aligned left)
	<option  <% if fieldtype(t)="snd" then response.write "selected" %> value="snd">embedded sound
	<option  <% if fieldtype(t)="tbx" then response.write "selected" %> value="tbx">generic text box
	<option  <% if fieldtype(t)="rtx" then response.write "selected" %> value="rtx">text box (aligned right)
	<option  <% if fieldtype(t)="ltx" then response.write "selected" %> value="ltx">text box (aligned left)
	<option  <% if fieldtype(t)="tit" then response.write "selected" %> value="tit">title 
	<option  <% if fieldtype(t)="sig" then response.write "selected" %> value="sig">signer
	<option  <% if fieldtype(t)="qut" then response.write "selected" %> value="qut">question title
	<option  <% if fieldtype(t)="qus" then response.write "selected" %> value="qus">question small
	<option  <% if fieldtype(t)="qul" then response.write "selected" %> value="qul">question large
	<option  <% if fieldtype(t)="ans" then response.write "selected" %> value="ans">answer
	<option  <% if fieldtype(t)="dal" then response.write "selected" %> value="dal">date (long)
	<option  <% if fieldtype(t)="dat" then response.write "selected" %> value="dat">date (short)
	<option  <% if fieldtype(t)="lhd" then response.write "selected" %> value="lhd">long header
	<option  <% if fieldtype(t)="shd" then response.write "selected" %> value="shd">short header
	</select></td><td>&nbsp;</td><tr>
<%
end sub

function storelines(inner, number)
'writes time remaining for a specific action using the FSO
	set timefile=fso.opentextfile(server.mappath(toolpath & "/" & tool & ".tol"), 1, true)
	if not timefile.atendofstream then
		getums=timefile.readall
	end if
	timefile.close
    getums=cleanup(getums)
	times=split(getums, chr(13))
	set timefile=fso.createtextfile(server.mappath(toolpath & "/"  & tool & ".tol"),  true)
	for xx=0 to 19
	if xx <= ubound(times) then 
			'response.write(xx & "<br>")
		if xx=number then
			timefile.writeline(inner)
		else 
			timefile.writeline(times(xx))
		end if
	end if
	next
	timefile.close
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
%>

<!--#include virtual="/include/CCFooter.asp"-->
