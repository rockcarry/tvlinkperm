<!-- #include file ="common.asp" -->

<%
    if Request.Cookies(strAdminLoginPwd) <> strAdminLoginPwd then
        Response.Redirect(strLoginPageName)
        Response.End()
    end if
%>

<%
    dim tabVisitRuleTable(8)
    tabVisitRuleTable(0) = "tabVisitRuleTable"
    tabVisitRuleTable(1) = "50%"
    tabVisitRuleTable(2) = strAdminPageName
    tabVisitRuleTable(3) = "submit.asp"
    tabVisitRuleTable(4) = "���"
    tabVisitRuleTable(5) = "IP ��ַ"
    tabVisitRuleTable(6) = "MAC ��ַ"
    tabVisitRuleTable(7) = "����Ȩ��"
    tabVisitRuleTable(8) = "����"

    dim tabVisitRecordTable(11)
    tabVisitRecordTable(0)  = "tabVisitRuleRecord"
    tabVisitRecordTable(1)  = "80%"
    tabVisitRecordTable(2)  = strAdminPageName
    tabVisitRecordTable(3)  = "submit.asp"
    tabVisitRecordTable(4)  = "���"
    tabVisitRecordTable(5)  = "IP ��ַ"
    tabVisitRecordTable(6)  = "MAC ��ַ"
    tabVisitRecordTable(7)  = "���ʼ���"
    tabVisitRecordTable(8)  = "������"
    tabVisitRecordTable(9)  = "����Ȩ��"
    tabVisitRecordTable(10) = "λ����Ϣ"
    tabVisitRecordTable(11) = "����"

    function MakePageTableItemAdminStr(name, id)
        dim str
        select case name
        case tabVisitRuleTable(0)
            str = "<a href=""submit.asp?optr=" & strOptrDeleteVisitRule
            str = str & "&id=" & id & """>ɾ��</a>&nbsp"
            str = str & "<a href=""submit.asp?optr=" & strOptrModifyVisitRulePage
            str = str & "&id=" & id & """>�޸�</a>"
        case tabVisitRecordTable(0)
            str = "<a href=""submit.asp?optr=" & strOptrDeleteVisitRecord
            str = str & "&id=" & id & """>ɾ��</a>"
        end select
        MakePageTableItemAdminStr = str
    end function

    function MakePageLinkString(table, page, link)
        MakePageLinkString = "<a href=""" & table(3) & "?optr=" & strOptrTablePageSubmit & "&"
        MakePageLinkString = MakePageLinkString & "name=" & table(0) & "&page=" & page & "&"
        MakePageLinkString = MakePageLinkString & "disp=" & table(2)
        MakePageLinkString = MakePageLinkString & """>" & link & "</a>"
    end function

    function CheckMACPermitted(mac)
        dim rs, sql
        set rs = Server.CreateObject("ADODB.recordset")
        sql = "SELECT * FROM PermittedMACTable WHERE MAC='" & mac & "'"
        rs.Open sql, conn
        if rs.EOF then
            CheckMACPermitted = false
        else
            CheckMACPermitted = true
        end if
    end function

    'table(0) - name
    'table(1) - width
    'table(2) - display page
    'table(3) - submit page
    'table(4) - title
    sub DisplayTableByPage(table, sql)
        dim rs, x, i, page, color

        set rs = Server.CreateObject("ADODB.recordset")
        rs.Open sql, conn, 1

        Response.Write("<table border=""1"" width=""" & table(1) & """>" & vbcrlf)
        Response.Write("<tr>")
        for i=4 to ubound(table)
            Response.Write("<th>" & table(i) & "</th>")
        next
        Response.Write("</tr>" & vbcrlf)

        rs.PageSize = nTablePageSize
        page = Request.Cookies(table(0))("page")
        if page = "" then page = "0"
        page = cint(page)
        if page > rs.PageCount then page = rs.PageCount
        if page < 1 then page = 1
        if not rs.EOF then
            rs.AbsolutePage = page
        end if

        for i=1 to rs.PageSize
            if not rs.EOF then
                if table(0) = tabVisitRecordTable(0) then
                    color = ""
                    if nVisitRecordQueryCond = 4 then
                        color = "bgcolor=""#ff8888"""
                    else
                        if not CheckMACPermitted(rs("MAC")) then
                            color = " bgcolor=""#ff8888"""
                        end if
                    end if
                end if

                Response.Write("<tr align=""center""" & color & ">" & vbcrlf)
                for each x in rs.Fields
                    Response.Write("<td>" & x.value & "</td>")
                next
                Response.Write("<td>" & MakePageTableItemAdminStr(table(0), rs("ID")) & "</td>")
                Response.Write("</tr>" & vbcrlf)
                rs.MoveNext()
            else
                exit for
            end if
        next

        Response.Write("</table>" & vbcrlf)
        Response.Write("total: " & rs.RecordCount & " ")
        Response.Write("page: " & page & "/" & rs.PageCount & " ")
        Response.Write(MakePageLinkString(table, 1,            " ��ҳ "))
        Response.Write(MakePageLinkString(table, page - 1,     " ��ҳ "))
        Response.Write(MakePageLinkString(table, page + 1,     " ��ҳ "))
        Response.Write(MakePageLinkString(table, rs.PageCount, " βҳ "))
        for i=page-5 to page+5
            if i >= 1 and i <= rs.PageCount then
                Response.Write(MakePageLinkString(table, i, " " & i & " "))
            end if
        next

        rs.Close()
        set rs = nothing
    end sub

    dim strVisitRecordQueryText(5)
    strVisitRecordQueryText(0) = "ȫ�����ʼ�¼"
    strVisitRecordQueryText(1) = "���ڷ��ʼ�¼"
    strVisitRecordQueryText(2) = "������ʼ�¼"
    strVisitRecordQueryText(3) = "���շ��ʼ�¼"
    strVisitRecordQueryText(4) = "δ��ȨMAC��¼"
    strVisitRecordQueryText(5) = "����ȨMAC��¼"

    dim nVisitRecordQueryCond
    nVisitRecordQueryCond = Request.Cookies(tabVisitRecordTable(0))("cond")
    if nVisitRecordQueryCond = "" then
        nVisitRecordQueryCond = 0
    else
        nVisitRecordQueryCond = cint(nVisitRecordQueryCond)
    end if

    sub DisplayVisitRecordQueryOption()
        dim str, i
        Response.Write("<select name=""cond"">" & vbcrlf)
        for i=0 to ubound(strVisitRecordQueryText)
            str = "<option value=""" & i & """"
            if i = nVisitRecordQueryCond then str = str & "selected"
            str = str & ">" & strVisitRecordQueryText(i) & "</option>" & vbcrlf
            Response.Write(str)
        next
        Response.Write("</select>" & vbcrlf)
    end sub
%>

<% OpenConn() %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0
  Transitional//EN" "http://wDisplayTableByPageww.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
  <meta http-equiv="Content-Language" content="zh-cn" />
  <title>����ҳ��</title>
</head>

<body>
<h1>����ҳ��</h1>
&nbsp;<a href="submit.asp?optr=<%=strOptrAdminLogout%>">[�˳�����]</a>
&nbsp;<a href="help.html">[ʹ��˵��]</a>
<!--
&nbsp;<a href="initdb.asp?optr=<%=strOptrResetDatabase%>">[�������ݿ�]</a>
-->
&nbsp;<a href="initdb.asp?optr=<%=strOptrImportIPTable%>">[����IP��ַ��]</a>
&nbsp;<a href="initdb.asp?optr=<%=strOptrImportMACTable%>">[����MAC��Ȩ��]</a>
<hr/>

<h2>���ʹ���</h2>
<% DisplayTableByPage tabVisitRuleTable, "SELECT * FROM VisitRuleTable" %>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrAddVisitRule%>" />
  IP: <input type="text" name="ip" value="*" />
  MAC:<input type="text" name="mac"value="*" />
  <input type="radio"  name="perm" value="1" />allowed
  <input type="radio"  name="perm" value="0" checked="checked" />forbidden
  <input type="submit" value="��ӹ���" />
</form>

<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrClearVisitRule%>" />
  <input type="submit" value="��շ��ʹ���" />
</form>
<br/>

<h2>���ʼ�¼</h2>
<table>
<tr>
<td>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr"  value="<%=strOptrVisitRecordCond%>" />
  <input type="hidden" name="table" value="<%=tabVisitRecordTable(0)%>" />
  <% DisplayVisitRecordQueryOption() %>
  <input type="submit" value="��ѯ" />
</form>
</td>
<td>&nbsp;</td>
<td>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr"  value="<%=strOptrVisitRecordCond%>" />
  <input type="hidden" name="table" value="<%=tabVisitRecordTable(0)%>" />
  <% if nVisitRecordQueryCond = 6 then%>
  <input type="radio" name="cond" value="6" checked="checked" />IP
  <input type="radio" name="cond" value="7" />MAC
  <% else %>
  <input type="radio" name="cond" value="6" />IP
  <input type="radio" name="cond" value="7" checked="checked" />MAC
  <% end if %>
  <input type="text"  name="ipmac" value="<%=Request.Cookies(tabVisitRecordTable(0))("ipmac")%>"/>
  <input type="submit" value="��ѯ" />
</form>
</td>
</tr>
</table>

<%
    dim strVisitRecordQuerySQL(7)
    strVisitRecordQuerySQL(0)  = "SELECT * FROM VisitRecordTable"
    strVisitRecordQuerySQL(1)  = "SELECT * FROM VisitRecordTable WHERE VisitLocation = 'China'"
    strVisitRecordQuerySQL(2)  = "SELECT * FROM VisitRecordTable WHERE VisitLocation <> 'China'"
if nDataBaseType = 1 or nDataBaseType = 2 then
    strVisitRecordQuerySQL(3)  = "SELECT * FROM VisitRecordTable WHERE DateDiff('d', VisitLastTime, Date()) = 0"
else
    strVisitRecordQuerySQL(3)  = "SELECT * FROM VisitRecordTable WHERE DateDiff('d', VisitLastTime, GetDate()) = 0"
end if
    strVisitRecordQuerySQL(4)  = "SELECT * FROM VisitRecordTable WHERE NOT EXISTS (SELECT NULL FROM PermittedMACTable WHERE VisitRecordTable.MAC=PermittedMACTable.MAC)"
    strVisitRecordQuerySQL(5)  = "SELECT * FROM VisitRecordTable WHERE EXISTS (SELECT NULL FROM PermittedMACTable WHERE VisitRecordTable.MAC=PermittedMACTable.MAC)"
    strVisitRecordQuerySQL(6)  = "SELECT * FROM VisitRecordTable WHERE IP  = '" & Request.Cookies(tabVisitRecordTable(0))("ipmac") & "'"
    strVisitRecordQuerySQL(7)  = "SELECT * FROM VisitRecordTable WHERE MAC = '" & Request.Cookies(tabVisitRecordTable(0))("ipmac") & "'"

    DisplayTableByPage tabVisitRecordTable, strVisitRecordQuerySQL(nVisitRecordQueryCond)
%>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr"  value="<%=strOptrClearVisitRecord%>" />
  <input type="hidden" name="cond"  value="<%=nVisitRecordQueryCond%>" />
  <input type="hidden" name="ipmac" value="<%=Request.Cookies(tabVisitRecordTable(0))("ipmac")%>"/>
  <input type="submit" value="��շ��ʼ�¼" />
</form>

</body>
</html>

<% CloseConn() %>
