<%
'CopyBot by Gus Mueller
'November 4 1998
'improvements ongoing

'on error resume next
localcachesize=33 'depth of check for duplicates when doing multiple file copies
cachesize=20     'the size of the handy cache pulldown menu
gifcachesize=20     'depth of check for duplicates when also copying included images
'for copy overs not using dev:
'transroot="c:\inetpub\wwwroot\update\trans\"
'cachepath="c:\inetpub\wwwroot\update\cache.txt"
grandioseoverride=request("gor")      'allows power users to replace the big & critical places
devname=request("devname")            'name of Development Server
bad=replace(request("bad"), " ", "")  'numbers of bad WWW servers
badbackup=replace(request("badbackup"), " ", "")   'to keep track of changes to the bad server list
puller=Request("puller")              'pulldown menu files (unused)
gifs=Request("gifs")                  'parse for gifs
pather=replace(Request("path"), " ", "")         'direct file/folder paths
cacher=Request("cacher")              'cached file paths (unused)
schedfile=Request("schedfile")        'schedule files
username=Request("username")          'user name
foljer=Request("foljer")              'direct folder paths
delfoljer=Request("delfoljer")        'delete folder paths
delet=Request("delet")                'deletion paths 
typer=request("type")                 'the nature of the copy
serverstart=request("serverstart")    'lowest live server to copy to

'let us no longer allow different varieties of simultaneous copy-over
'thus this is a kind of hierarchy
if pather<>"" or foljer<>"" then 
	schedfile=""
	delfoljer=""
	delet=""
end if
if pather<>"" then
	foljer=""
end if
if schedfile<>"" then
	delfoljer=""
	delet=""
end if
if delet<>"" then
	delfoljer=""
end if
server.scripttimeout=9999
set fsc=createobject("scripting.filesystemobject")

'provision for the button copy overs that makes it completely radical
if schedfile<>"" then 
	if instr(schedfile, "routine") then
		typer="stl"
	else
		typer="dtl"
	end if
end if

%>
<!--#include file=drives.asp-->
<%
'response.write(dev & " " & dev2 & "<br>")
if serverstart="" then serverstart=0
dontdeletebadschedule=0
badstartsaving=false
redim localcache(localcachesize)
redim gifcache(gifcachesize)
redim getumser(cachesize+2)

'update list of bad servers both here and on dev
if bad<>badbackup then
	set worst=fsc.createtextfile(dev & "\ccroot\update\bad.txt",  true)
	worst.writeline(bad)
	worst.close
	fsc.copyfile dev & "\ccroot\update\bad.txt", server.mappath("\update\bad.txt"), true
end if
%>
<html>
<head>
<title>the CopyBot</title>
</head>
<body bgcolor=ff9999 text=black><p><font face=arial>
<h3>This is where the files get copied from 
<%
'set up the process for various kinds of copy overs
if typer="dts" then 
	startsaving=false
	servers=0
	response.write(" DEV to Staging")
elseif typer="stl" then
	serverstart=1
	dev=w(0)
	response.write(" Staging to WWWx")
elseif typer="std" then
	startsaving=false
	tempstore=dev
	dev=w(0)
	w(0)=tempstore
	servers=0
	response.write(" Staging to DEV")
elseif typer="dt2" then
	startsaving=false
	servers=0
	w(0)=dev2
	response.write(" DEV to DEV2")
elseif typer="2dt" then
	startsaving=false
	servers=0
	w(0)=dev
	dev=dev2
	response.write(" DEV2 to DEV")
else
	response.write(" DEV to Staging and WWWx")
end if
if typer="dtl" or typer="stl" then 
	majoraction=true
else
	majoraction=false
end if
filecount=0
deleted=0
%>
</h3>
The name of the file or folder is mentioned first, then the mapped drives being copied to (as it happens).<p></font><font face=courier size=1>
<hr size=1 noshade>
<%
if username<>"" then
	'start saving a transcript if there are bad servers	
	retrystring="0"
	if schedfile<>"" then
		if instr(schedfile, "baddrive")>0 then
			if startsaving=true then
				badstartsaving=true
			end if
			startsaving=false
			retrystring="retry"
		end if
	end if
	
	'set up the storing of paths for servers that are currently bad
	if startsaving or badstartsaving then
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
			mi = "0" & mi
		end if
		se=cstr(second(time))
		if len(se)=1 then 
			se = "0" & se
		end if
		'determine if there is a baddrive file so we can append to it if there is
		set fold=fsc.getfolder(transroot)
		set filies=fold.files
		baddrivedatepart=""
		for each fy in filies
			'response.write(fy.name & "<br>")
			if asc(left(fy.name,1))<58 and asc(left(fy.name,1))>47 and instr(fy.name, "sync")=0 then
				if instr(fy.name, "baddrive") and not instr(fy.name, "retry") then
					if not badstartsaving then
						baddrivedatepart=left(fy.name, 4)
						badplace=transroot & baddrivedatepart & "baddrive"  & ".tra"
					end if
				end if
			end if
		next
		if  baddrivedatepart="" or badstartsaving then
			'make a new baddrive file
			if badstartsaving then
				badplace=transroot & mo & da & "baddrive" & retrystring & ".tra"
			else
			   badplace=transroot & mo & da & "baddrive" & ".tra"
			end if
		end if
	end if
	
	'clean up pathers now!!
		if pather<>"" then
			pather=replace(pather, "/", "\")
			pather=trim(pather)
			if left(pather,1) <> "\" then 
				pather="\" & pather
			end if 
		end if
	logaction("-------" & username &"  "& date() &"  "& time() & "-------")
	'update the handy cache
	if pather<>"" and (fsc.fileexists(dev&pather) or fsc.folderexists(dev&pather))  then
	set cac=fsc.opentextfile(cachepath, 1, true)
	poscount=0
	t=0
		do until cac.atendofstream
			getums=trim(replace(replace(trim(cac.readline), chr(10), ""), chr(13), ""))
			if left(getums, 1)="\" then
				if lcase(pather)=lcase(getums) then
					poscount=poscount+1
				else
					getumser(t)=getums
					t=t+1
				end if 'pather=getums
			end if 'is a path
		loop
	cac.close
	set cac=fsc.opentextfile(cachepath, 2, true)
	rend=t-1
	if poscount=0 then
		if left(pather,1)="\" then
			cac.writeline(pather)
		end if
			rend=rend-1
		end if 'poscount=0
		if rend<cachesize and poscount>0 then
			cac.writeline(pather)
		end if
		if rend>cachesize then
			rend=rend-1
		end if
		for r=0 to rend
			cac.writeline(getumser(r))
		next
		if rend<cachesize and poscount=0 then
			cac.writeline(pather)
		end if
		cac.close
	end if
	
	'deal with pathers that might be folders
	if right(pather, 1)="\" then
		pather=left(pather, len(pather)-1)
	end if
	if fsc.folderexists(dev & pather) then
		foljer=pather
		pather=""
	end if
	
	'check if there is a schedulefile
	if schedfile<>"" then
		if fsc.fileexists(transroot & schedfile) then
		set sched=fsc.opentextfile(transroot & schedfile)
			do until sched.atendofstream
			getums=trim(sched.readline)
			getums=replace(replace(getums, "/", "\"), " ", "")
				if getums<>"" then
					if left(getums,1)="\" then
					'determine if this copy over is only to certain servers that were bad
					questpos=instr(getums, "?")
						if questpos<>0 then
						parts=split(getums, "?")
						urlpart=parts(0)
						badpart=parts(1)
						copybad=split(badpart, ",")
						copybadtop=ubound(copybad)
						getums=urlpart
						end if
						if instr(getums, ".")=0 and right(getums, 1)<>"\" then
						getums=getums & "\"
						end if
					overwrought=0
						for localscan=0 to localcachesize 
							if localcache(localscan)=getums & badpart then
							overwrought=overwrought+1
							end if
						next
					if overwrought=0 then
						for localscan = 0 to localcachesize-1
						localcache(localscan)=localcache(localscan+1)
						next
					localcache(localcachesize)=getums & badpart
					response.write(getums & " ")
					filetrue=1
						if right(trim(getums), 1)="\" then
						filetrue=0
						getums=left(getums, len(getums)-1)
							if grandiose(getums)=1 then 
							response.write("<font color=ffffff>:too grandiose!</font><br>" & chr(13))
							else
							if fsc.folderexists(dev & getums) then
								if badpart="" then
								response.write(dirmaker(getums))
									for specificserver= serverstart to servers
										if w(specificserver)<>"" then
										fsc.copyfolder dev & getums, w(specificserver) & getums, true
											'if fsc.folderexists(w(specificserver) & getums) then
												'attribs= fsc.getfolder(w(specificserver) & getums).attributes
												'if attribs and 1 then
													'attribs=attribs-1
												'end if
												'fsc.getfolder(w(specificserver) & getums).attributes = attribs
											'end if
											if checkfoldercount(getums, specificserver) then
												response.write(" " & specificserver)
											else
												response.write(" " & "<i><b><font color=00cc00>" & specificserver & "</font></b></i>")
											end if
										else
										response.write(" <font color=00ffff>" & specificserver & "</font>")
										end if
									next
								logaction("folder: " & getums)
								else
								'copy over to missed drives
								if grandiose(getums)=1 then 
									response.write("<font color=ffffff>:too grandiose!</font><br>" & chr(13))
								else
									if fsc.folderexists(dev & getums) then
								    baddirmaker(getums)
									for specificserver= 0 to copybadtop
										if  w(cint(copybad(specificserver)))<>"" then
										fsc.copyfolder dev & getums, w(cint(copybad(specificserver))) & getums, true
											if fsc.folderexists(w(cint(copybad(specificserver))) & getums) then
												attribs= fsc.getfolder(w(cint(copybad(specificserver))) & getums).attributes
												if attribs and 1 then
													attribs=attribs-1
												end if
												'fsc.getfolder(w(cint(copybad(specificserver)))& getums).attributes = attribs
											end if																			
										if checkfoldercount(getums, cint(copybad(specificserver))) then
											response.write(" " & cint(copybad(specificserver)))
										else
											response.write(" " & "<i><b><font color=00cc00>" & cint(copybad(specificserver)) & "</font></b></i>")
										end if		
										else
										response.write(" <font color=99ffff>" & cint(copybad(specificserver)) & "</font>")
										dontdeletebadschedule=dontdeletebadschedule+1
										end if
									next
									else 
								    response.write("<font color=ffffff>:The folder does not exist!</font><br>" & chr(13))
									end if
								end if
								end if
								badsaver(getums&"?"&bad)
								filecount=filecount+1
								response.write("<br>" & chr(13))
							else
								response.write("<font color=ffffff>:The folder does not exist!</font><br>" & chr(13))
							end if ' grandiose check
						end if ' ends with a slash
					end if ' is a folder
					if filetrue=1 then
						if fsc.fileexists(dev & getums) then
							if grandiose(getums)=1 then 
								response.write("<font color=ffffff>:too grandiose!</font><br>" & chr(13))
							else
								if badpart="" then
									response.write(dirmaker(getums))
									for specificserver= serverstart to servers
										if w(specificserver)<>"" then
											if fsc.fileexists(w(specificserver) & getums) then
												attribs= fsc.getfile(w(specificserver) & getums).attributes
												if attribs and 1 then
													attribs=attribs-1
												end if
												fsc.getfile(w(specificserver) & getums).attributes = attribs
											end if
										'response.write(w(specificserver) & getums & "<br>")
											fsc.copyfile dev & getums, w(specificserver) & getums, true
											if fsc.fileexists(w(specificserver) & getums) then
												attribs= fsc.getfile(w(specificserver) & getums).attributes
												if attribs and 1 then
													attribs=attribs-1
												end if
												fsc.getfile(w(specificserver) & getums).attributes = attribs
											end if
											if checkfilesize(getums, specificserver) then
												response.write(" " & specificserver)
											else
												response.write(" " & "<i><b><font color=00cc00>" & specificserver & "</font></b></i>")
											end if
										else
											response.write(" <font color=00ffff>" & specificserver & "</font>")
										end if
									next
									logaction (getums)
								else
								'copy over to missed drives
									if fsc.fileexists(dev & getums) then
									'response.write("<br>*"& dev & getums & "<br>*")
									baddirmaker(getums)
										for specificserver= 0 to copybadtop
											if w(cint(copybad(specificserver)))<>"" then
												if fsc.fileexists(w(cint(copybad(specificserver))) & getums) then
													attribs= fsc.getfile(w(cint(copybad(specificserver))) & getums).attributes
													if attribs and 1 then
														attribs=attribs-1
													end if
													fsc.getfile(w(cint(copybad(specificserver))) & getums).attributes = attribs
												end if
											fsc.copyfile dev & getums, w(cint(copybad(specificserver))) & getums, true
	
											if checkfilesize(getums, cint(copybad(specificserver))) then
												response.write(" " & cint(copybad(specificserver)))
											else
												response.write(" " & "<i><b><font color=00cc00>" & cint(copybad(specificserver)) & "</font></b></i> ")
											end if
											else
											response.write(" <font color=99ffff>" & cint(copybad(specificserver)) & "</font> ")
											dontdeletebadschedule=dontdeletebadschedule+1
											end if
										next
									else 
									response.write("<font color=ffffff>:The file does not exist!</font><br>" & chr(13))
									end if 'file exists
								end if 'grandiose check
								if gifs="yes" then
									response.write(getgifs(getums))
								end if 'is gifs set
								badsaver (getums&"?"&bad)
								filecount=filecount+1
								response.write("<br>" & chr(13))
							end if
						else
						response.write("<font color=ffffff>:The file does not exist!</font><br>" & chr(13))
						end if 'file exists
					end if 'is a file as opposed to a folder
					end if 'overwrought test
				end if 'is actually a path
			end if 'is not empty
		loop
	sched.close
	'delete schedule file when it's no longer useful
	'but only if the copy over was a dramatic one and, if there was one, the
	'badschedule copy over was completely successful
	if (typer="dtl" or typer="stl") and instr(schedfile, "routine")=0 then
	fsc.deletefile transroot & schedfile
	end if
	end if 'the schedfile exists
	end if 'there is a schedfile
	
	'check if there is a absolute file path
	if pather<>"" then
	response.write("<br><b>" & pather & "</b> ")
		if grandiose(pather)=1 then 
		response.write("<font color=ffffff>:too grandiose!</font><br>" & chr(13))
		else
			if fsc.fileexists(dev & pather) then
			response.write(dirmaker(pather))
				for specificserver= serverstart to servers
					if w(specificserver)<>"" then
						if fsc.fileexists(w(specificserver) & pather) then
							attribs= fsc.getfile(w(specificserver) & pather).attributes
							if attribs and 1 then
								attribs=attribs-1
							end if
							fsc.getfile(w(specificserver) & pather).attributes = attribs
						end if
						fsc.copyfile dev & pather, w(specificserver) & pather, true
						if fsc.fileexists(w(specificserver) & pather) then
							attribs= fsc.getfile(w(specificserver) & pather).attributes
							if attribs and 1 then
								attribs=attribs-1
							end if
							fsc.getfile(w(specificserver) & pather).attributes = attribs
						end if
						if checkfilesize(pather, specificserver) then
							response.write(" " & specificserver)
						else
							response.write(" " & "<i><b><font color=00cc00>" & specificserver & "</font></b></i>")
						end if
					else
						response.write(" <font color=00ffff>" & specificserver & "</font>")
					end if
				next
					logaction(pather)
					badsaver (pather&"?"&bad)
				if gifs="yes" then
					response.write(getgifs(pather))
				end if 'is gifs set
				filecount=filecount+1
			else
			response.write("<font color=ffffff>:The file/folder does not exist!</font><br>" & chr(13))
			end if 'pather path exists
		end if 'grandiose check
	end if 'there is a pather
	
	'check if there is an absolute folder path
	if foljer<>"" then
		foljer=replace(foljer, "/" , "\")
		foljer=trim(foljer)
	if right(foljer, 1) = "\" then
		foljer=left(foljer, len(foljer)-1)
	end if
	if left(foljer,1) <> "\" then 
		foljer="\" & foljer
	end if 
	response.write("<br><b>" & foljer & "</b> ")
		if grandiose(foljer)=1 then 
			response.write("<font color=ffffff>:too grandiose!</font><br>" & chr(13))
		else
			if fsc.folderexists(dev & foljer) then
				response.write(dirmaker(foljer))
				for specificserver= serverstart to servers
					if w(specificserver)<>"" then
						'response.write("<br>" & w(specificserver) & foljer & "<br>")
						fsc.copyfolder dev & foljer, w(specificserver) & foljer, true
						if fsc.folderexists(w(specificserver) & foljer) then
							attribs= fsc.getfolder(w(specificserver) & foljer).attributes
							if attribs and 1 then
								attribs=attribs-1
							end if
							'fsc.getfolder(w(specificserver) & foljer).attributes = attribs
						end if					
						if checkfoldercount(foljer, specificserver) then
							response.write(" " & specificserver)
					else
						response.write(" " & "<i><b><font color=00cc00>" & specificserver & "</font></b></i>")
					end if
					else
						response.write(" <font color=00ffff>" & specificserver & "</font>")
					end if
				next
					filecount=filecount+1
					logaction("folder:" & foljer)
					badsaver (foljer&"?"&bad)
			else
				response.write("<font color=ffffff>:The folder does not exist!</font><br>" & chr(13))
			end if 'folder exists
		end if 'grandiose check
	end if 'there is a folder copy request
	
	'check if there is a delete order
	if delet<>"" then
	delet=trim(replace(delet, "/" , "\"))
	if left(delet,1) <> "\" then 
	delet="\" & delet
	end if 
	response.write("<br><font color=990000>deleting:</font> <b>" & delet & "</b> ")
		for specificserver= serverstart to servers
			if w(specificserver)<>"" then
				if fsc.fileexists(w(specificserver) & delet) then
					if fsc.getfile(w(specificserver) & delet).attributes and 2 then
						fsc.getfile(w(specificserver) & delet).attributes = fsc.getfile(w(specificserver) & delet).attributes-2
					end if
					fsc.deletefile w(specificserver) & delet
					response.write(" " & specificserver)
				else
						response.write(" <font color=ff3333>" & specificserver & "</font>")
				end if
			else 
				response.write(" <font color=00ffff>" & specificserver & "</font>")
			end if
		next
	deleted=deleted+1
			logaction("deleted: " & delet)
	end if 'there is a delete order request
	
	'check if there is a delete folder order
	if delfoljer<>"" then
		response.write("<br><font color=990000>deleting:</font> <b>" & delfoljer & "</b> ")
		if grandiose(delfoljer)=1 then 
			response.write("<font color=ffffff>:too grandiose!</font><br>" & chr(13))
		else
			for specificserver= serverstart to servers
			
				if fsc.folderexists(w(specificserver) & delfoljer) then
					if w(specificserver)<>"" then
						fsc.deletefolder w(specificserver) & delfoljer
						response.write(" " & specificserver)
					else
						response.write(" <font color=00ffff>" & specificserver & "</font>")
					end if
				else 
					response.write(" <font color=ff3333>" & specificserver & "</font>")
				end if
			next
		deleted=deleted+1
				logaction("deleted: " & delfoljer)
		end if 'grandiose check
	end if 'there is a delete order request
	
	%>
	<hr size=1 noshade>
	<%
	if filecount=0 then
		response.write("no files were copied over because there were no orders to do so")
	else
		response.write("<br>" & filecount & " files or folders were copied over.")
	end if
	if deleted>0 then
		response.write("<br>" & deleted & " files or folders were deleted.")
	end if

	'Extra Action
	if instr(schedfile, "today")<>0 then%>
		<h2>Extra Action: Today in CC Default Update</h2>
		<!--#include file="today.asp"-->
		<p>
	<%elseif instr(schedfile, "campus")<>0 then%>
		<h2>Extra Action: Campus Image Update</h2>
		<!--#include file="campusimage.asp"-->
		<p>
	<%end if 'extra actions necessary
	if badstartsaving then
		if fsc.fileexists( transroot & mo & da & "baddrive" & ".tra") then
			set oldbad=fsc.opentextfile(transroot & mo & da & "baddrive" & ".tra", 8)
			set newbad=fsc.opentextfile(transroot & mo & da & "baddrive" & retrystring & ".tra", 1)
				do until newbad.atendofstream
					getums3=newbad.readline
					oldbad.writeline(getums3)
				loop 
			oldbad.close
			newbad.close
			fsc.deletefile(transroot & mo & da & "baddrive" & retrystring & ".tra")
		else
			fsc.movefile transroot & mo & da & "baddrive" & retrystring & ".tra", transroot & mo & da & "baddrive" & ".tra"
		end if
	end if
else 
	response.write("<center><h1><font color=red>INTRUDER DETECTED!!</h1><br><img src=http://spies.com/~gus/dan.gif border=10></font><br><font face=arial size=6>nothing was done!</font></center><pre>")
	response.end
end if 'user name check

function grandiose(pat)
'this functions determines if folder or file
'creations or replacements 
'are too grandiose
pat=lcase(pat)
	if pat="\ccroot" or pat="\ccgifs" or pat="\test" or pat="\ccroot\ads" or pat="\ccroot\cclub" or pat="\ccroot\members" or pat="\ccroot\campus" or pat="\ccgifs\campus\content" or pat="\ccgifs\campus\" or pat="\" or instr(pat, "*") or instr(pat, "?") or pat="\ccroot\include\defaultinclude.asp" or pat="\ccroot\include\default.asp" or pat="\ccroot\school"  then
		if grandioseoverride<>"1" then
		grandiose=1
		else
		grandiose=0
		end if
	else 
	grandiose=0
	end if
end function

function getgifs(pat)
	'this function parses out gif paths from HTML files and copies them to live
	suffix=lcase(right(trim(pat), 3))
	if suffix<>"asp" and suffix<>"htm" and suffix <> "tml" and suffix<>"tpw" then
		getgifs="<font color=ffffff>Not an HTML file</font><br>"
	else
		response.write("<br>")
		set gimp=fsc.opentextfile(dev&pat)
		do until gimp.atendofstream
			mine=lcase(gimp.readline)
			if len(mine)>5 then
				if instr(lcase(mine), "src=") then
					start=1
					nextplace=instr(start, mine, "src=")
					do until nextplace=0
						if len(mine)>5 and posish < len(mine) and start< len(mine) then
						'response.write("<pre>" & mine & "</pre>")
							posish=instr(start, mine, "src=")
						'response.write("<pre>" & mine & "</pre>")
							posish2=instr(posish, mine, " ")
						end if
						if posish > 0 then
							if len(mine)>posish+4 and posish2-(posih+4)>0 then
								urler=trim(mid(mine, posish+4, (posish2)-(posish+4)))
								urler=trim(replace(urler, chr(34), ""))
								urler=replace(urler, "/" , "\")
								grafsuf=lcase(right(urler, 3))
								if grafsuf= "gif" or grafsuf ="jpg" then
									gifover=0
									for gifscan=0 to gifcachesize 
										if gifcache(gifscan)=urler then
											gifover=gifover+1
										end if
									next
									if gifover=0 then
										for gifscan= 0 to gifcachesize-1
											gifcache(gifscan)=gifcache(gifscan+1)
										next
										gifcache(gifcachesize)=urler
										response.write("<br>" & urler & chr(13))
										response.write(dirmaker(urler))
										if badpart="" then
											for specificserver= serverstart to servers
												if w(specificserver)<>"" then
													fsc.copyfile dev & urler, w(specificserver) & urler, true
													response.write(" " & specificserver)
												else
													response.write(" <font color=00ffff>" & specificserver & "</font>")
												end if
											next
											response.write("<br>")
										else
											'copy over to missed drives
											for specificserver= 0 to copybadtop
												if  w(cint(copybad(specificserver)))<>"" then
													fsc.copyfile dev & urler, w(cint(copybad(specificserver))) & urler, true
													response.write(" " & cint(copybad(specificserver)))
												else
													response.write(" <font color=99ffff>" & cint(copybad(specificserver)) & "</font>")
													dontdeletebadschedule=dontdeletebadschedule+1
												end if
											next
											response.write("<br>")
										end if
										logaction(urler)
										if startsaving then
											badsaver (urler&"?"&bad)
										end if
										filecount=filecount+1
									end if
								end if 'graphic sufix is good
							end if
						end if 'position is good
						start=posish2
						if len(mine)>5 and start < len(mine) then
						'response.write("<pre>" & mine & "</pre>")
							nextplace=instr(start, lcase(mine), "src=")
						else 
							nextplace=0
						end if
					loop
				end if  'line contains src=
			end if
		loop
		getgifs=""
	end if 'is actually a html file
	response.write("<br>")
end function

'this function creates folders as needed on the live servers
function dirmaker(gifpath) 
'makes new folders when needed
start=1
posish=0
do while instr(start, gifpath, "\") <>0
	prelimer=left(gifpath, instr(start, gifpath, "\")) 
	if right(prelimer, 1)="\" then
		prelimer=left(prelimer, len(prelimer)-1)
	end if
	'response.write("<br>"& prelimer & "-" & posish)
	if grandiose(prelimer)<>1 and prelimer<>("") then
		needforafolder=0
		for specificserver=serverstart to servers
			if w(specificserver)<>"" then
				if not fsc.folderexists(w(specificserver) & prelimer) then needforafolder=needforafolder+1
			end if
		next
		if needforafolder = 0 then
			response.write("")
		else
			response.write("<br><b>Creating Folder: <br>" & prelimer & ":</b> ")
			for specificserver=serverstart to servers
				if w(specificserver)<>"" then
					if not fsc.folderexists(w(specificserver) & prelimer) then
						fsc.createfolder(w(specificserver) & prelimer)
						response.write( specificserver & " ")
					else
						response.write(" <font color=ff0000>" & specificserver & "</font> ")
					end if
				else
					response.write(" <font color=00ffff>" & specificserver & "</font> ")
				end if
			next
			response.write("<br>")
		end if 
	end if 'grandiose check
	posish=instr(start, gifpath, "\")
	start=posish+1
loop
dirmaker=""
end function

function baddirmaker(gifpath) 
'makes new folders when needed
start=1
posish=0
do while instr(start, gifpath, "\") <>0
	prelimer=left(gifpath, instr(start, gifpath, "\")) 
	if right(prelimer, 1)="\" then
		prelimer=left(prelimer, len(prelimer)-1)
	end if
	if grandiose(prelimer)<>1 and prelimer<>("") then
		needforafolder=0
		for specificserver=0 to copybadtop
			if w(cint(copybad(specificserver))) <>"" then
				if not fsc.folderexists(w(cint(copybad(specificserver))) & prelimer) then needforafolder=needforafolder+1
			end if
		next
		if needforafolder = 0 then
			response.write("")
		else
			response.write("<br><b>Creating Folder: <br>" & prelimer & ":</b> ")
			for specificserver = 0 to copybadtop
				if w(cint(copybad(specificserver)))<>"" then
					if not fsc.folderexists(w(cint(copybad(specificserver))) & prelimer) then
						fsc.createfolder(w(cint(copybad(specificserver))) & prelimer)
						response.write( cint(copybad(specificserver)) & " ")
					else 
						response.write(" <font color=ff0000>" & cint(copybad(specificserver)) & "</font> ")
					end if
				else
					response.write(" <font color=00ffff>" & cint(copybad(specificserver)) & "</font> ")
				end if
			next
			response.write("<br>")
		end if 
	end if 'grandiose check
	posish=instr(start, gifpath, "\")
	start=posish+1
loop
baddirmaker=""
end function


function checkfilesize(filepath, webserver)
'performs a test to be sure the copied file on live is the same size as the source file on dev
devsize=fsc.getfile(dev & filepath).size
websize=fsc.getfile(w(webserver) & filepath).size
if devsize<>websize then
checkfilesize=false
else
checkfilesize=true
end if
end function

function checkfoldercount(filepath, webserver)
'performs a test to be sure the copied folder on live contains the same number of items as 
'the counterpart on dev
devsize=fsc.getfolder(dev & filepath).files.count + fsc.getfolder(dev & filepath).subfolders.count
websize=fsc.getfolder(w(webserver) & filepath).files.count + fsc.getfolder(w(webserver) & filepath).subfolders.count
if devsize<>websize then
checkfoldercount=false
else
checkfoldercount=true
end if
end function

function logaction(logpath)
'logs all multi-server web copies
if majoraction then
	set snooplog=fsc.opentextfile(logfile, 8, true)
	snooplog.writeline(logpath)
	snooplog.close
end if
end function

function badsaver(badpath)
'logs all copies that should have gone to servers removed from production
if startsaving or badstartsaving then
	set badsave=fsc.opentextfile(badplace, 8, true)
	badsave.writeline(badpath)
	badsave.close
end if
end function
%>
<hr size=1 noshade>
<font face=arial>
All major transactions made here were logged to a file at <a href=http://content/log.txt>logfile</a>.  <br>
If you go there, be sure to hit refresh.
</font>
<font face=arial size=3>
<p>Return to the <a href=http://devoffice.office.xxxxx.com/update/cop.asp>Copy Over Page</a></font>
</body>
</html>
