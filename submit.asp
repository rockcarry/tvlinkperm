<!-- #include file ="common.asp" -->
<% OpenConn() %>

<%
    dim bSubmitSucceed
    dim strErrorMessage
    dim strRedirectTo

    bSubmitSucceed  = true
    strErrorMessage = ""
    strRedirectTo   = strAdminPageName

    dim strVisitRecordDelSQL(7)
    strVisitRecordDelSQL(0) = "DELETE FROM VisitRecordTable"
    strVisitRecordDelSQL(1) = "DELETE FROM VisitRecordTable WHERE VisitLocation = 'China'"
    strVisitRecordDelSQL(2) = "DELETE FROM VisitRecordTable WHERE VisitLocation <> 'China'"
if nDataBaseType = 1 or nDataBaseType = 2 then
    strVisitRecordDelSQL(3) = "DELETE FROM VisitRecordTable WHERE DateDiff('d', VisitLastTime, Date()) = 0"
else
    strVisitRecordDelSQL(3) = "DELETE FROM VisitRecordTable WHERE DateDiff('d', VisitLastTime, GetDate()) = 0"
end if
    strVisitRecordDelSQL(4) = "DELETE FROM VisitRecordTable WHERE NOT EXISTS (SELECT NULL FROM PermittedMACTable WHERE VisitRecordTable.MAC=PermittedMACTable.MAC)"
    strVisitRecordDelSQL(5) = "DELETE FROM VisitRecordTable WHERE EXISTS (SELECT NULL FROM PermittedMACTable WHERE VisitRecordTable.MAC=PermittedMACTable.MAC)"
    strVisitRecordDelSQL(6) = "DELETE FROM VisitRecordTable WHERE IP  = '" & Request.Form("ipmac") & "'"
    strVisitRecordDelSQL(7) = "DELETE FROM VisitRecordTable WHERE MAC = '" & Request.Form("ipmac") & "'"

    strOptrCur = Request("optr")
    select case strOptrCur
    case strOptrAdminLogin
        CheckAdminLogin()
    case strOptrAdminLogout
        DoAdminLogout()
    case strOptrAddVisitRule
        AddVisitRule()
    case strOptrDeleteVisitRule
        DelVisitRule()
    case strOptrModifyVisitRulePage
        ModifyVisitRulePage()
    case strOptrModifyVisitRuleDoIt
        ModifyVisitRuleDoIt()
    case strOptrClearVisitRule
        ClearVisitRule()
    case strOptrDeleteVisitRecord
        DelVisitRecord()
    case strOptrVisitRecordCond
        VisitRecordCondQuery()
    case strOptrClearVisitRecord
        ClearVisitRecord()
    case strOptrExportVisitRecord
        ExportVisitRecord()
    case strOptrTablePageSubmit
        HandleTablePageSubmit()
    end select
%>

<%
    sub CheckAdminLogin()
        if Request.Form("pwd") = strAdminLoginPwd then
            Response.Cookies(strAdminLoginPwd) = strAdminLoginPwd
            Response.Cookies(strAdminLoginPwd).Expires = (now() + 1)
        end if
    end sub

    sub DoAdminLogout()
        Response.Cookies(strAdminLoginPwd) = ""
        Response.Cookies(strAdminLoginPwd).Expires = (now() - 1)
    end sub

    sub AddVisitRule()
        dim ip, mac, remark, perm, sql, rs
        ip     = Request.Form("ip"    )
        mac    = Request.Form("mac"   )
        remark = Request.Form("remark")
        perm   = Request.Form("perm"  )
        mac    = lcase(mac)
        remark = left(remark, 64)

        if ip = "" OR mac = "" then
            strErrorMessage = "IP 和 MAC 地址不能为空！<br/>"
            strErrorMessage = strErrorMessage & "<a href=""" & strAdminPageName & """>返回</a>"
            bSubmitSucceed  = false
            exit sub
        end if

        set rs = Server.CreateObject("ADODB.recordset")
        sql = "SELECT * FROM VisitRuleTable WHERE IP = '" & ip & "'" & " AND MAC = '" & mac & "'"
        rs.Open sql, conn, 1, 3

        if not rs.EOF then
            strErrorMessage = "该 IP 和 MAC 地址的访问规则已经存在！<br/>"
            strErrorMessage = strErrorMessage & "<a href=""" & strAdminPageName & """>返回</a>"
            bSubmitSucceed  = false
        else
            rs.AddNew()
            rs("IP" )    = ip
            rs("MAC")    = mac
            rs("Remark") = remark
            rs("VisitPermission") = perm
            rs.Update()
        end if

        rs.Close()
        set rs = nothing
    end sub

    sub DelVisitRule()
        dim id, sql
        id  = Request.QueryString("id")
        sql = "DELETE FROM VisitRuleTable WHERE ID = " & id
        conn.Execute(sql)
    end sub

    dim nModVisitRuleID
    dim strModVisitRuleIP
    dim strModVisitRuleMAC
    dim strModVisitRuleRemark
    dim nModVisitRulePerm
    sub ModifyVisitRulePage()
        dim sql, rs
        set rs = Server.CreateObject("ADODB.recordset")
        sql = "SELECT * FROM VisitRuleTable WHERE ID = " & Request.QueryString("id")
        rs.Open sql, conn
        if not rs.EOF then
            nModVisitRuleID      = rs("ID")
            strModVisitRuleIP    = rs("IP")
            strModVisitRuleMAC   = rs("MAC")
            strModVisitRuleRemark= rs("Remark")
            nModVisitRulePerm    = rs("VisitPermission")
        else
            strErrorMessage = "要修改的 IP 和 MAC 地址的访问规则不存在！<br/>"
            strErrorMessage = strErrorMessage & "<a href=""" & STR_ADMIN_PAGE_NAME & """>返回</a>"
            bSubmitSucceed  = false
        end if
        rs.Close()
        set rs = nothing
    end sub

    sub ModifyVisitRuleDoIt()
        dim id, sql, rs
        set rs = Server.CreateObject("ADODB.recordset")
        id  = Request.Form("id")
        sql = "SELECT * FROM VisitRuleTable WHERE ID = " & id
        rs.Open sql, conn, 1, 3
        if not rs.EOF then
            rs("Remark"         ) = Request.Form("remark")
            rs("VisitPermission") = Request.Form("perm"  )
            rs.Update()
        end if
        rs.Close()
        set rs = nothing
    end sub

    sub ClearVisitRule()
        conn.Execute("DELETE FROM VisitRuleTable")
    end sub

    sub DelVisitRecord()
        conn.Execute("DELETE FROM VisitRecordTable WHERE ID = " & Request.QueryString("id"))
    end sub

    sub VisitRecordCondQuery()
        dim table, cond, ipmac
        table = Request.Form("table")
        cond  = Request.Form("cond" )
        ipmac = Request.Form("ipmac")
        Response.Cookies(table)("cond" ) = cond
        Response.Cookies(table)("ipmac") = ipmac
    end sub

    sub ClearVisitRecord()
        dim sql, cond
        cond = cint(Request.Form("cond"))
        sql  = strVisitRecordDelSQL(cond)
        conn.Execute(sql)
    end sub

    sub ExportVisitRecord()
        dim fs, txt, rs, sql, x, line

        set fs = Server.CreateObject("Scripting.FileSystemObject")
        set txt= fs.OpenTextFile(Server.MapPath(strExpVisitRecord), 2, true)
        set rs = Server.CreateObject("ADODB.recordset")

        sql = "SELECT * FROM VisitRecordTable"
        rs.Open sql, conn

        do while not rs.EOF
            line = ""
            for each x in rs.Fields
                line = line & x.value & chr(9)
            next
            txt.WriteLine(line)
            rs.MoveNext()
        loop

        rs.Close()
        txt.Close()
        set rs = nothing
        set txt= nothing
        set fs = nothing

        strRedirectTo = strExpVisitRecord
    end sub

    sub HandleTablePageSubmit()
        dim name, page, disp
        name = Request.QueryString("name")
        page = Request.QueryString("page")
        disp = Request.QueryString("disp")
        Response.Cookies(name)("page") = page
    end sub
%>

<% CloseConn() %>

<% if strOptrCur = strOptrModifyVisitRulePage then %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0
  Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
  <meta http-equiv="Content-Language" content="zh-cn" />
  <title>管理页面</title>
</head>
<body>

<h2>访问规则</h2>

<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrModifyVisitRuleDoIt%>" />
  <input type="hidden" name="id"   value="<%=nModVisitRuleID%>" />
  <table>
    <tr><td>IP:</td><td><%=strModVisitRuleIP %></td></tr>
    <tr><td>MAC:</td><td><%=strModVisitRuleMAC%></td></tr>
    <tr><td>备注:</td><td><input type="text" name="remark" value="<%=strModVisitRuleRemark%>" size="64" /></td></tr>
    <tr>
      <td>权限:</td>
      <td>
      <% if nModVisitRulePerm = 1 then %>
        <input type="radio" name="perm" value="1" checked="checked" /> allowed
        <input type="radio" name="perm" value="0" /> forbidden
      <% else %>
        <input type="radio" name="perm" value="1" /> allowed
        <input type="radio" name="perm" value="0" checked="checked" /> forbidden
      <% end if %>
      </td>
    </tr>
    <tr><td></td><td><input type="submit" value="修改访问规则" /></td></tr>
  </table>
</form>

</body>
</html>
<% else %>

<%
    if bSubmitSucceed then
        Response.Redirect(strRedirectTo)
    else
        Response.Write(strErrorMessage)
    end if
%>

<% end if %>
