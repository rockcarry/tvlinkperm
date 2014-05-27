<!-- #include file ="common.asp" -->

<%
    if Request.Cookies(strAdminLoginPwd) <> strAdminLoginPwd then
        Response.Redirect(strLoginPageName)
        Response.End()
    end if
%>

<%
    dim tabVisitRuleTable(9)
    tabVisitRuleTable(0) = "tabVisitRuleTable"
    tabVisitRuleTable(1) = "80%"
    tabVisitRuleTable(2) = strAdminPageName
    tabVisitRuleTable(3) = "submit.asp"
    tabVisitRuleTable(4) = "���"
    tabVisitRuleTable(5) = "IP ��ַ"
    tabVisitRuleTable(6) = "MAC ��ַ"
    tabVisitRuleTable(7) = "��ע"
    tabVisitRuleTable(8) = "����Ȩ��"
    tabVisitRuleTable(9) = "����"

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
        for i=page-10 to page+10
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

    sub DisplayQueryOptions(name, opts, sel)
        dim str, i
        Response.Write("<select name=""" & name & """>" & vbcrlf)
        for i=0 to ubound(opts)
            str = "<option value=""" & i & """"
            if i = sel then str = str & "selected"
            str = str & ">" & opts(i) & "</option>" & vbcrlf
            Response.Write(str)
        next
        Response.Write("</select>" & vbcrlf)
    end sub

    function GetDistinctMACNum(cond)
        dim rs, sql
        set rs = Server.CreateObject("ADODB.recordset")
        sql = "SELECT DISTINCT MAC FROM VisitRecordTable" & cond
        rs.Open sql, conn, 1
        GetDistinctMACNum = rs.RecordCount
        rs.Close()
        set rs = nothing
    end function
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
<% if nWriteOutType = 3 then %>
&nbsp;<a href="initdb.asp?optr=<%=strOptrImportChannelData%>">[����Ƶ������]</a>
<% end if %>
<hr/>

<h2>���ʹ���</h2>
<% DisplayTableByPage tabVisitRuleTable, "SELECT * FROM VisitRuleTable" %>
<br/><br/>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrAddVisitRule%>" />
  <table>
    <tr><td>IP:</td><td><input type="text" name="ip" value="*" size="18" /></td></tr>
    <tr><td>MAC:</td><td><input type="text" name="mac" value="*" size="18" /></td></tr>
    <tr><td>��ע:</td><td><input type="text" name="remark" size="64" /></td></tr>
    <tr>
      <td>Ȩ��:</td>
      <td>
        <input type="radio"  name="perm" value="1" />allowed
        <input type="radio"  name="perm" value="0" checked="checked" />forbidden
      </td>
    </tr>
    <tr><td></td><td><input type="submit" value="��ӷ��ʹ���" /></td></tr>
  </table>
</form>


<h2>���ʼ�¼</h2>

<%
    dim strQueryCondsVisitTime(4)
    dim strQueryCondsVisitPerm(2)
    dim strQueryCondsMACPerm(2)
    dim strQueryCondsSortType(4)
    dim strQueryCondCountryCode
    dim strQueryCondIPValue
    dim strQueryCondMACValue
    dim nQueryCondVisitTimeValue
    dim nQueryCondVisitPermValue
    dim nQueryCondMACPermValue
    dim nQueryCondSortTypeValue
    dim strSQLVisitTime(4)
    dim strSQLVisitPerm(2)
    dim strSQLMACPerm(2)
    dim strSQLSortType(4)
    dim strSQLCondStr

    strQueryCondsVisitTime(0) = "*"
    strQueryCondsVisitTime(1) = "һ����"
    strQueryCondsVisitTime(2) = "һ����"
    strQueryCondsVisitTime(3) = "һ����"
    strQueryCondsVisitTime(4) = "һ����"

    strQueryCondsVisitPerm(0) = "*"
    strQueryCondsVisitPerm(1) = "��Ȩ��"
    strQueryCondsVisitPerm(2) = "��Ȩ��"

    strQueryCondsMACPerm(0)   = "*"
    strQueryCondsMACPerm(1)   = "����Ȩ"
    strQueryCondsMACPerm(2)   = "δ��Ȩ"

    strQueryCondsSortType(0)   = "Ĭ�Ϸ�ʽ"
    strQueryCondsSortType(1)   = "����ʱ��+"
    strQueryCondsSortType(2)   = "����ʱ��-"
    strQueryCondsSortType(3)   = "���ʼ���+"
    strQueryCondsSortType(4)   = "���ʼ���-"

    strQueryCondCountryCode  = Request.Cookies("query_cond")("country_code")
    strQueryCondIPValue      = Request.Cookies("query_cond")("ip_value")
    strQueryCondMACValue     = Request.Cookies("query_cond")("mac_value")
    nQueryCondVisitTimeValue = Request.Cookies("query_cond")("visit_time")
    nQueryCondVisitPermValue = Request.Cookies("query_cond")("visit_perm")
    nQueryCondMACPermValue   = Request.Cookies("query_cond")("mac_perm")
    nQueryCondSortTypeValue  = Request.Cookies("query_cond")("sort_type")

    if strQueryCondCountryCode  = "" then strQueryCondCountryCode  = "*"
    if strQueryCondIPValue      = "" then strQueryCondIPValue      = "*"
    if strQueryCondMACValue     = "" then strQueryCondMACValue     = "*"
    if nQueryCondVisitTimeValue = "" then nQueryCondVisitTimeValue = 0
    if nQueryCondVisitPermValue = "" then nQueryCondVisitPermValue = 0
    if nQueryCondMACPermValue   = "" then nQueryCondMACPermValue   = 0
    if nQueryCondSortTypeValue  = "" then nQueryCondSortTypeValue  = 0
    nQueryCondVisitTimeValue = cint(nQueryCondVisitTimeValue)
    nQueryCondVisitPermValue = cint(nQueryCondVisitPermValue)
    nQueryCondMACPermValue   = cint(nQueryCondMACPermValue  )
    nQueryCondSortTypeValue  = cint(nQueryCondSortTypeValue )

    strSQLVisitTime(0) = ""
if nDataBaseType = 1 or nDataBaseType = 2 then
    strSQLVisitTime(1) = " AND DateDiff('d', VisitLastTime, Date()) = 0"
    strSQLVisitTime(2) = " AND DateDiff('w', VisitLastTime, Date()) = 0"
    strSQLVisitTime(3) = " AND DateDiff('m', VisitLastTime, Date()) = 0"
    strSQLVisitTime(4) = " AND DateDiff('yyyy', VisitLastTime, Date()) = 0"
else
    strSQLVisitTime(1) = " AND DateDiff(day, VisitLastTime, GetDate()) = 0"
    strSQLVisitTime(2) = " AND DateDiff(week, VisitLastTime, GetDate()) = 0"
    strSQLVisitTime(3) = " AND DateDiff(month, VisitLastTime, GetDate()) = 0"
    strSQLVisitTime(4) = " AND DateDiff(year, VisitLastTime, GetDate()) = 0"
end if

    strSQLVisitPerm(0) = ""
    strSQLVisitPerm(1) = " AND VisitPermission=1"
    strSQLVisitPerm(2) = " AND VisitPermission=0"

    strSQLMACPerm(0)   = ""
    strSQLMACPerm(1)   = " AND EXISTS (SELECT NULL FROM PermittedMACTable WHERE VisitRecordTable.MAC=PermittedMACTable.MAC)"
    strSQLMACPerm(2)   = " AND NOT EXISTS (SELECT NULL FROM PermittedMACTable WHERE VisitRecordTable.MAC=PermittedMACTable.MAC)"

    strSQLSortType(0)  = ""
    strSQLSortType(1)  = " ORDER BY VisitLastTime"
    strSQLSortType(2)  = " ORDER BY VisitLastTime DESC"
    strSQLSortType(3)  = " ORDER BY VisitCounter"
    strSQLSortType(4)  = " ORDER BY VisitCounter DESC"

    strSQLCondStr = "WHERE VisitLocation='" & strQueryCondCountryCode & "'"
    strSQLCondStr = strSQLCondStr & " AND IP='" & strQueryCondIPValue & "'"
    strSQLCondStr = strSQLCondStr & " AND MAC='" & strQueryCondMACValue & "'"
    strSQLCondStr = strSQLCondStr & "MAC='" & strQueryCondMACValue & "' AND "

    function MakeQueryStr0(section, str)
        if str <> "*" then
            if left(str, 1) = "!" then
                MakeQueryStr0 = " AND " & section & "<>'" & mid(str, 2) & "'"
            else
                MakeQueryStr0 = " AND " & section & "='" & str & "'"
            end if
        else
            MakeQueryStr0 = ""
        end if
    end function

    strSQLCondStr = " WHERE 1=1"
    strSQLCondStr = strSQLCondStr & MakeQueryStr0("VisitLocation", strQueryCondCountryCode)
    strSQLCondStr = strSQLCondStr & MakeQueryStr0("IP" , strQueryCondIPValue )
    strSQLCondStr = strSQLCondStr & MakeQueryStr0("MAC", strQueryCondMACValue)
    strSQLCondStr = strSQLCondStr & strSQLVisitTime(nQueryCondVisitTimeValue)
    strSQLCondStr = strSQLCondStr & strSQLVisitPerm(nQueryCondVisitPermValue)
    strSQLCondStr = strSQLCondStr & strSQLMACPerm(nQueryCondMACPermValue)
%>

<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrVisitRecordCond%>" />
  <table>
    <tr><td>���Ҵ���</td><td>IP</td><td>MAC</td><td>����ʱ��</td><td>����Ȩ��</td><td>MAC��Ȩ</td><td>����ʽ</td><td></td></tr>
    <tr>
      <td><input name="country_code" type="text" value="<%=strQueryCondCountryCode%>" size="8" /></td>
      <td><input name="ip_value"     type="text" value="<%=strQueryCondIPValue    %>" size="17"/></td>
      <td><input name="mac_value"    type="text" value="<%=strQueryCondMACValue   %>" size="17"/></td>
      <td><% DisplayQueryOptions "visit_time", strQueryCondsVisitTime, nQueryCondVisitTimeValue %></td>
      <td><% DisplayQueryOptions "visit_perm", strQueryCondsVisitPerm, nQueryCondVisitPermValue %></td>
      <td><% DisplayQueryOptions "mac_perm"  , strQueryCondsMACPerm  , nQueryCondMACPermValue   %></td>
      <td><% DisplayQueryOptions "sort_type" , strQueryCondsSortType , nQueryCondSortTypeValue  %></td>
      <td><input type="submit" value="��ѯ"/></td>
    </tr>
  </table>
</form>

<%
    dim strSQLVisitRecord
    strSQLVisitRecord = "SELECT * FROM VisitRecordTable"
    strSQLVisitRecord = strSQLVisitRecord & strSQLCondStr & strSQLSortType(nQueryCondSortTypeValue)
    DisplayTableByPage tabVisitRecordTable, strSQLVisitRecord
%>
<br/>�ܹ��� <%=GetDistinctMACNum(strSQLCondStr)%> ����ͬ�� MAC.<br/></br>
<table>
<tr>
<td>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrClearVisitRecord%>" />
  <input type="hidden" name="cond" value="<%=strSQLCondStr%>" />
  <input type="submit" value="ɾ�����ʼ�¼" />
</form>
</td>
<td>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrExportVisitRecord%>" />
  <input type="hidden" name="cond" value="<%=strSQLCondStr%>" />
  <input type="submit" value="�������ʼ�¼" />
</form>
</td>
<td>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrExportMACPermTable%>" />
  <input type="submit" value="����MAC��Ȩ��" />
</form>
</td>
</tr>
</table>

</body>
</html>

<% CloseConn() %>
