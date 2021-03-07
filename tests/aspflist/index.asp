<%@ Language=VBScript%>
<%
'List the contents of the current
'***** Enter a BOOKMARK TITLE for this web page in the line below ***
strTitle = "Title goes here"
'***** Enter a MAIN HEADING for this web page in the line below *****
strHeading = "Heading goes here"
'You don't need to change anything below this line
'********************************************************************
%>
<HTML>
<HEAD>
<%response.write ("<TITLE>" & strTitle & "</TITLE>")%>
</HEAD>
<BODY>
    <%
    response.write ("<H1>" & strHeading & "</H1>")
    set FileSysObj=CreateObject("Scripting.FileSystemObject")
    strFileAndPath = request.servervariables("SCRIPT_NAME")
    strPathOnly = Mid(strFileAndPath,1 ,InStrRev(strFileAndPath, "/"))
    strFullPath = server.mappath(strPathOnly)
    set fldr=FileSysObj.GetFolder(strFullPath)
    response.write("<H2>Folders list</H2>")
    set FolderList = fldr.SubFolders
    For Each FolderIndex in FolderList
        Response.Write("<A HREF='" & FolderIndex.name & "'>" & FolderIndex.name & "</A><BR>")
    Next
    response.write("<H2>Files list</H2>")
    set FileList = fldr.Files
    For Each FileIndex in FileList
        'This bit excludes this page (and other asp files) from the list of links
        if Lcase(right(FileIndex.Name, 4)) <> ".asp" then
            Response.Write("<A HREF='" & FileIndex.name & "'>" & FileIndex.name & "</A><BR>")
        end if
    Next
    %>
</BODY>
</HTML>