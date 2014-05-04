<!-- #include file ="common.asp" -->

<% OpenConn() %>

<%
    if strOptrCur = "" then strOptrCur = strOptrResetDatabase

    sub ImportIPTableFromTxt(url)
        dim fs, txt, line, items, sql, rs
        set fs = Server.CreateObject("Scripting.FileSystemObject")
        set txt= fs.OpenTextFile(Server.MapPath(url))
        set rs = Server.CreateObject("ADODB.recordset")
        on error resume next
        conn.Execute("DELETE FROM IPLocationTable")
        sql = "SELECT * FROM IPLocationTable"
        rs.Open sql, conn, 1, 3
        do while not txt.AtEndOfStream
            line = txt.ReadLine()
            if line <> "" then
                items= split(line, chr(9))
                rs.AddNew()
                rs("NUM_IP_START") = cdbl(items(0))
                rs("NUM_IP_END"  ) = cdbl(items(1))
                rs("STR_IP_START") = items(2)
                rs("STR_IP_END"  ) = items(3)
                rs("LOCATION"    ) = items(4)
                rs.Update()
            end if
        loop
        rs.Close()
        txt.Close()
        set rs = nothing
        set txt= nothing
        set fs = nothing
    end sub

    sub ImportMACTableFromTxt(url)
        dim fs, txt, line, sql, rs
        set fs = Server.CreateObject("Scripting.FileSystemObject")
        set txt= fs.OpenTextFile(Server.MapPath(url))
        set rs = Server.CreateObject("ADODB.recordset")
        on error resume next
        conn.Execute("DELETE FROM PermittedMACTable")
        sql = "SELECT * FROM PermittedMACTable"
        rs.Open sql, conn, 1, 3
        do while not txt.AtEndOfStream
            line = trim(txt.ReadLine())
            if line <> "" then
                rs.AddNew()
                rs("MAC") = line
                rs.Update()
            end if
        loop
        rs.Close()
        txt.Close()
        set rs = nothing
        set txt= nothing
        set fs = nothing
    end sub

    sub CreateSystemTables()
        dim sql(3), x
        if nDataBaseType = 1 or nDataBaseType = 2 then
            sql(0) = "CREATE TABLE IPLocationTable(ID autoincrement primary key, NUM_IP_START float, NUM_IP_END float, STR_IP_START text(16), STR_IP_END text(16), LOCATION text(16))"
            sql(1) = "CREATE TABLE VisitRuleTable(ID autoincrement primary key, IP text(16), MAC text(18), VisitPermission int)"
            sql(2) = "CREATE TABLE VisitRecordTable(ID autoincrement primary key, IP text(16), MAC text(18), VisitCounter int, VisitLastTime datetime, VisitPermission int, VisitLocation text(16))"
            sql(3) = "CREATE TABLE PermittedMACTable(ID autoincrement primary key, MAC text(18))"
        end if
        if nDataBaseType = 3 or nDataBaseType = 4 then
            sql(0) = "CREATE TABLE IPLocationTable(ID int identity primary key, NUM_IP_START float, NUM_IP_END float, STR_IP_START nvarchar(16), STR_IP_END nvarchar(16), LOCATION nvarchar(16))"
            sql(1) = "CREATE TABLE VisitRuleTable(ID int identity primary key, IP nvarchar(16), MAC nvarchar(18), VisitPermission int)"
            sql(2) = "CREATE TABLE VisitRecordTable(ID int identity primary key, IP nvarchar(16), MAC nvarchar(18), VisitCounter int, VisitLastTime datetime, VisitPermission int, VisitLocation nvarchar(16))"
            sql(3) = "CREATE TABLE PermittedMACTable(ID int identity primary key, MAC nvarchar(18))"
        end if
        for each x in sql
            conn.Execute(x)
        next
        conn.Execute("CREATE INDEX IPIndex ON IPLocationTable(NUM_IP_START, NUM_IP_END)")
        conn.Execute("CREATE INDEX MACIndex ON PermittedMACTable(MAC)")
    end sub

    sub DeleteSystemTables()
        dim sql(3), x
        sql(0) = "DROP TABLE IPLocationTable"
        sql(1) = "DROP TABLE VisitRuleTable"
        sql(2) = "DROP TABLE VisitRecordTable"
        sql(3) = "DROP TABLE PermittedMACTable"
        for each x in sql
            on error resume next
            conn.Execute(x)
            if err.number<>0 then
                err.Clear()
            end if
        next
    end sub


    sub ImportChannelData(url)
        dim ado
        set ado  = Server.CreateObject("Adodb.Stream")
        ado.Type = 1
        ado.Open()
        ado.LoadFromFile(Server.MapPath(url))
        application("ChannelFileData") = ado.Read
        application("ChannelFileSize") = ado.Size
        ado.Close()
        set ado = nothing
    end sub
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0
  Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
  <meta http-equiv="Content-Language" content="zh-cn" />
  <title>
  <% select case strOptrCur %>
  <% case strOptrResetDatabase  %>
    重置数据库
  <% case strOptrImportIPTable  %>
    导入IP地址库
  <% case strOptrImportMACTable %>
    导入MAC授权库
  <% end select %>
  </title>
</head>

<body>
<%
    if strOptrCur = strOptrResetDatabase then
        Response.Write("正在初始化数据库...<br/>" & vbcrlf)
        Response.Flush()
        DeleteSystemTables()
        CreateSystemTables()
    end if

    if strOptrCur = strOptrResetDatabase or strOptrCur = strOptrImportIPTable then
        Response.Write("正在导入 IP 地址库...<br/>" & vbcrlf)
        Response.Flush()
        ImportIPTableFromTxt(strIPLocTableTxt)
    end if

    if strOptrCur = strOptrResetDatabase or strOptrCur = strOptrImportMACTable then
        Response.Write("正在导入 MAC 授权库...<br/>" & vbcrlf)
        Response.Flush()
        ImportMACTableFromTxt(strMACTableTxt)
    end if

    if strOptrCur = strOptrResetDatabase or strOptrCur = strOptrImportChannelData then
        if nWriteOutType = 3 then
            Response.Write("正在导入频道数据文件...<br/>" & vbcrlf)
            Response.Flush()
            ImportChannelData(strChannelXmlPath)
        end if
    end if

    Response.Write("完成！<br/>" & vbcrlf)
    Response.Flush()
%>
<a href="<%=strAdminPageName%>">返回管理</a>
</body>

</html>

<% CloseConn() %>

