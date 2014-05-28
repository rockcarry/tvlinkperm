<!-- #include file ="common.asp" -->
<% OpenConn() %>

<%
    dim bSubmitSucceed
    dim strErrorMessage
    dim strRedirectTo

    bSubmitSucceed  = true
    strErrorMessage = ""
    strRedirectTo   = strAdminPageName

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
        ExportVisitRecord strExpVisitRecord, "SELECT * FROM VisitRecordTable" & Request.Form("cond")
    case strOptrExportMACPermTable
        ExportVisitRecord strExpMacPermTab, "SELECT * FROM PermittedMACTable"
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
        Response.Cookies("query_cond")("oneipmultimac")= Request.Form("oneipmultimac")
        Response.Cookies("query_cond")("country_code") = Request.Form("country_code")
        Response.Cookies("query_cond")("ip_value")     = Request.Form("ip_value")
        Response.Cookies("query_cond")("mac_value")    = Request.Form("mac_value")
        Response.Cookies("query_cond")("visit_time")   = Request.Form("visit_time")
        Response.Cookies("query_cond")("visit_perm")   = Request.Form("visit_perm")
        Response.Cookies("query_cond")("mac_perm")     = Request.Form("mac_perm")
        Response.Cookies("query_cond")("sort_type")    = Request.Form("sort_type")
    end sub

    sub ClearVisitRecord()
        conn.Execute("DELETE FROM VisitRecordTable" & Request.Form("cond"))
    end sub

    sub ExportVisitRecord(file, sql)
        dim fs, txt, rs, x, line

        set fs = Server.CreateObject("Scripting.FileSystemObject")
        set txt= fs.OpenTextFile(Server.MapPath(file), 2, true)
        set rs = Server.CreateObject("ADODB.recordset")

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

        strRedirectTo = file
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
