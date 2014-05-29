<!-- #include file ="common.asp" -->

<%
    if Request.Cookies(strAdminLoginPwd) <> strAdminLoginPwd then
        Response.Redirect(strLoginPageName)
        Response.End()
    end if
%>

<%
    dim tabVisitRule(8)
    tabVisitRule(0) = "tabVisitRule"
    tabVisitRule(1) = "90%"
    tabVisitRule(2) = "submit.asp"
    tabVisitRule(3) = "���"
    tabVisitRule(4) = "IP ��ַ"
    tabVisitRule(5) = "MAC ��ַ"
    tabVisitRule(6) = "��ע"
    tabVisitRule(7) = "����Ȩ��"
    tabVisitRule(8) = "����"

    dim tabVisitRecord(11)
    tabVisitRecord(0)  = "tabVisitRecord"
    tabVisitRecord(1)  = "90%"
    tabVisitRecord(2)  = "submit.asp"
    tabVisitRecord(3)  = "���"
    tabVisitRecord(4)  = "IP ��ַ"
    tabVisitRecord(5)  = "MAC ��ַ"
    tabVisitRecord(6)  = "���ʼ���"
    tabVisitRecord(7)  = "������"
    tabVisitRecord(8)  = "����Ȩ��"
    tabVisitRecord(9)  = "λ����Ϣ"
    tabVisitRecord(10) = "MAC ��Ȩ"
    tabVisitRecord(11) = "����"

    dim tabOneIPMultiMac(5)
    tabOneIPMultiMac(0) = "tabOneIPMultiMac"
    tabOneIPMultiMac(1) = "90%"
    tabOneIPMultiMac(2) = "submit.asp"
    tabOneIPMultiMac(3) = "IP"
    tabOneIPMultiMac(4) = "MAC �ܸ���"
    tabOneIPMultiMac(5) = "�����ܼ���"

    function MakePageTableItemAdminStr(name, id)
        dim str
        select case name
        case tabVisitRule(0)
            str = "<a href=""submit.asp?optr=" & strOptrDeleteVisitRule
            str = str & "&id=" & id & """>ɾ��</a>&nbsp"
            str = str & "<a href=""submit.asp?optr=" & strOptrModifyVisitRulePage
            str = str & "&id=" & id & """>�޸�</a>"
        case tabVisitRecord(0)
            str = "<a href=""submit.asp?optr=" & strOptrDeleteVisitRecord
            str = str & "&id=" & id & """>ɾ��</a>"
        end select
        MakePageTableItemAdminStr = str
    end function

    function MakePageLinkString(table, page, link, flag)
        if flag then link = "<u>" & link & "</u>"
        MakePageLinkString = "<a href=""" & table(2) & "?optr=" & strOptrTablePageSubmit & "&"
        MakePageLinkString = MakePageLinkString & "name=" & table(0) & "&page=" & page
        MakePageLinkString = MakePageLinkString & """>" & link & "</a>&nbsp;"
    end function

    'table(0) - name
    'table(1) - width
    'table(2) - submit page
    'table(3) - title
    sub DisplayTableByPage(table, sql)
        dim rs, x, i, min, max, page, color, str

        set rs = Server.CreateObject("ADODB.recordset")
        rs.Open sql, conn, 1

        Response.Write("<table id=""datatab"" width=""" & table(1) & """>" & "<tr>")
        for i=3 to ubound(table)
            Response.Write("<th>" & table(i) & "</th>")
        next
        Response.Write("</tr>" & vbcrlf)

        rs.PageSize = nTablePageSize
        page = Request.Cookies(table(0))("page")
        if page = "" then page = 0 else page = cint(page)
        if page > rs.PageCount then page = rs.PageCount
        if page < 1 then page = 1
        if not rs.EOF then rs.AbsolutePage = page

        for i=1 to rs.PageSize
            if not rs.EOF then
                if (i mod 2) = 1 then color = " class=""alt""" else color = ""

                if table(0) = tabVisitRecord(0) then
                    if isnull(rs(7)) then color = " class=""warn"""
                end if

                Response.Write("<tr" & color & ">")
                for each x in rs.Fields
                    Response.Write("<td>" & x.value & "</td>")
                next

                if table(0) <> tabOneIPMultiMac(0) then
                    Response.Write("<td>" & MakePageTableItemAdminStr(table(0), rs(0)) & "</td>")
                end if

                Response.Write("</tr>" & vbcrlf)
                rs.MoveNext()
            else
                exit for
            end if
        next
        Response.Write("</table>" & vbcrlf)

        min = page - 7
        if min < 1 then min = 1
        max = min + 14
        if max > rs.PageCount then
            max = rs.PageCount
            min = max - 14
            if min < 1 then min = 1
        end if

        Response.Write("<table><tr><td>")
        Response.Write("total:" & rs.RecordCount & " " & "page:" & page & "/" & rs.PageCount & " ")
        if max-min > 0 then
            str =       MakePageLinkString(table, 1,            "��ҳ", false)
            str = str & MakePageLinkString(table, page - 1,     "��ҳ", false)
            str = str & MakePageLinkString(table, page + 1,     "��ҳ", false)
            str = str & MakePageLinkString(table, rs.PageCount, "βҳ", false)
            Response.Write(str)

            for i=min to max
                Response.Write(MakePageLinkString(table, i, cstr(i), i=page))
            next
        end if
        Response.Write("</td>")

        Response.Write("<td>")
        if rs.PageCount > 15 then
            str = "<form action=""submit.asp"" method=""post"">"
            str = str & "<input name=""optr"" type=""hidden"" value=""" & strOptrTablePageSubmit & """/>"
            str = str & "<input name=""name"" type=""hidden"" value=""" & table(0) & """/>"
            str = str & "<input name=""page"" type=""text"" size=""3"" value=""" & page & """/>"
            str = str & "<input type=""submit"" value=""GO""/>"
            str = str & "</form>"
            Response.Write(str)
        end if
        Response.Write("</td></tr></table>")

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
  <link rel="stylesheet" type="text/css" href="tvlinkperm.css" />
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
<% DisplayTableByPage tabVisitRule, "SELECT * FROM VisitRuleTable" %>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrAddVisitRule%>" />
  <table>
    <tr><td>IP:</td><td><input type="text" name="ip" value="*" size="18" /></td></tr>
    <tr><td>MAC:</td><td><input type="text" name="mac" value="*" size="18" /></td></tr>
    <tr><td>��ע:</td><td><input type="text" name="remark" size="64" /></td></tr>
    <tr>
      <td>Ȩ��:</td>
      <td>
        <input type="radio" name="perm" value="1" />allowed
        <input type="radio" name="perm" value="0" checked="checked" />forbidden
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
    dim strQueryCondOneIPMultiMac
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

    strQueryCondsSortType(0)  = "Ĭ�Ϸ�ʽ"
    strQueryCondsSortType(1)  = "����ʱ��+"
    strQueryCondsSortType(2)  = "����ʱ��-"
    strQueryCondsSortType(3)  = "���ʼ���+"
    strQueryCondsSortType(4)  = "���ʼ���-"

    strQueryCondOneIPMultiMac= Request.Cookies("query_cond")("oneipmultimac")
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
    strSQLCondStr = strSQLCondStr & MakeQueryStr0("VisitRecordTable.VisitLocation", strQueryCondCountryCode)
    strSQLCondStr = strSQLCondStr & MakeQueryStr0("VisitRecordTable.IP" , strQueryCondIPValue )
    strSQLCondStr = strSQLCondStr & MakeQueryStr0("VisitRecordTable.MAC", strQueryCondMACValue)
    strSQLCondStr = strSQLCondStr & strSQLVisitTime(nQueryCondVisitTimeValue)
    strSQLCondStr = strSQLCondStr & strSQLVisitPerm(nQueryCondVisitPermValue)
    strSQLCondStr = strSQLCondStr & strSQLMACPerm(nQueryCondMACPermValue)
%>

<table>
  <tr><td>���Ҵ���</td><td>IP</td><td>MAC</td><td>����ʱ��</td><td>����Ȩ��</td><td>MAC��Ȩ</td><td>����ʽ</td><td></td><td></td></tr>
  <tr>
    <form action="submit.asp" method="post">
    <input type="hidden" name="optr" value="<%=strOptrVisitRecordCond%>" />
    <input type="hidden" name="oneipmultimac" value="0" />
    <input type="hidden" name="tabname" value="<%=tabVisitRecord(0)%>" />
    <td><input name="country_code" type="text" value="<%=strQueryCondCountryCode%>" size="8" /></td>
    <td><input name="ip_value"     type="text" value="<%=strQueryCondIPValue    %>" size="17"/></td>
    <td><input name="mac_value"    type="text" value="<%=strQueryCondMACValue   %>" size="17"/></td>
    <td><% DisplayQueryOptions "visit_time", strQueryCondsVisitTime, nQueryCondVisitTimeValue %></td>
    <td><% DisplayQueryOptions "visit_perm", strQueryCondsVisitPerm, nQueryCondVisitPermValue %></td>
    <td><% DisplayQueryOptions "mac_perm"  , strQueryCondsMACPerm  , nQueryCondMACPermValue   %></td>
    <td><% DisplayQueryOptions "sort_type" , strQueryCondsSortType , nQueryCondSortTypeValue  %></td>
    <td><input type="submit" value="��ѯ"/></td>
    </form>

    <form action="submit.asp" method="post">
    <input type="hidden" name="optr" value="<%=strOptrVisitRecordCond%>" />
    <input type="hidden" name="oneipmultimac" value="1" />
    <input type="hidden" name="tabname" value="<%=tabOneIPMultiMac(0)%>" />
    <td><input type="submit" value="��IP��MAC"/></td>
    </form>
  </tr>
</table>

<%
    dim strSQLVisitRecord
    strSQLVisitRecord = "SELECT VisitRecordTable.*, PermittedMACTable.ID FROM VisitRecordTable LEFT JOIN PermittedMACTable ON VisitRecordTable.MAC=PermittedMACTable.MAC"
    strSQLVisitRecord = strSQLVisitRecord & strSQLCondStr & strSQLSortType(nQueryCondSortTypeValue)

    if strQueryCondOneIPMultiMac <> "1" then
        DisplayTableByPage tabVisitRecord, strSQLVisitRecord
    else
        DisplayTableByPage tabOneIPMultiMac, "SELECT IP, count(MAC), sum(VisitCounter) FROM VisitRecordTable GROUP BY IP HAVING count(MAC)>1 ORDER BY count(MAC), sum(VisitCounter) DESC"
    end if
%>
�ܹ��� <%=GetDistinctMACNum(strSQLCondStr)%> ����ͬ�� MAC.

<table>
<tr>
<% if strQueryCondOneIPMultiMac <> "1" then %>
<td>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrClearVisitRecord%>" />
  <input type="hidden" name="cond" value="<%=strSQLCondStr%>" />
  <input type="submit" value="ɾ�����ʼ�¼" />
</form>
</td>
<% end if %>
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
<td>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrAutoPermitMAC%>" />
  <input type="submit" value="Ϊ�з���Ȩ�޵�δ��Ȩ��MAC��Ȩ" />
</form>
</td>
</tr>
</table>

</body>
</html>

<% CloseConn() %>
