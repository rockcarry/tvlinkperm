<%
    option explicit

    dim strChannelXmlPath
    dim strUSAChanXmlPath
    dim strIPLocTableTxt
    dim strMACTableTxt
    dim strAdminLoginPwd
    dim strAdminPageName
    dim strLoginPageName
    dim strExpVisitRecord
    dim strExpMacPermTab
    dim bSaveVisitRecord
    dim nDataBaseType
    dim nWriteOutType
    dim nTablePageSize
    dim strChinaCode
    dim strUSACode

    strChannelXmlPath = "channel.xml"
    strUSAChanXmlPath = "usachan.xml"
    strIPLocTableTxt  = "iptable.txt"
    strMACTableTxt    = "mactable.txt"
    strAdminLoginPwd  = "www.tvbox.com"
    strAdminPageName  = "admin.asp"
    strLoginPageName  = "login.asp"
    strExpVisitRecord = "export_visit_rule.txt"
    strExpMacPermTab  = "export_mac_permmit.txt"
    bSaveVisitRecord  = true
    nDataBaseType     = 2
    nWriteOutType     = 1
    nTablePageSize    = 10
    strChinaCode      = "CN"
    strUSACode        = "USA"

    dim strAccessDBFile
    dim strAccessDBPWD
    dim strSQLServerHost
    dim strSQLServerUSER
    dim strSQLServerPWD
    dim strSQLServerDBN

    'access
    strAccessDBFile  = "tvbox.mdb"
    strAccessDBPWD   = "www.tvbox.com"

    'sqlserver
    strSQLServerHost = "WIN-GDONNGTOTGF\LOOLBOX_4_29"
    strSQLServerUSER = "sa"
    strSQLServerPWD  = "Loolbox2014"
    strSQLServerDBN  = "looltv_content"

    dim strconn
    select case nDataBaseType
    case 1
        strconn = "dbq=" & Server.MapPath(strAccessDBFile) & ";"
        strconn = strconn & "driver={microsoft access driver (*.mdb)};"
        strconn = strconn & "pwd=" & strAccessDBPWD
    case 2
        strconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(strAccessDBFile) & ";"
        strconn = strconn & "Persist Security Info=false;Jet OLEDB:Database Password=" & strAccessDBPWD
    case 3
        strconn = "Driver={SQL Server};SERVER=" & strSQLServerHost & ";"
        strconn = strconn & "DATABASE=" & strSQLServerDBN & ";"
        strconn = strconn & "UID=" & strSQLServerUSER & ";"
        strconn = strconn & "PWD=" & strSQLServerPWD
    case 4
        strconn = "provider=sqloledb;"
        strconn = strconn & "data source=" & strSQLServerHost & ";"
        strconn = strconn & "initial catalog=" & strSQLServerDBN & ";"
        strconn = strconn & "user id=" & strSQLServerUSER & ";"
        strconn = strconn & "password=" & strSQLServerPWD
    end select

    '++数据库连接++'
    dim conn
    sub OpenConn()
        set conn = Server.CreateObject("ADODB.Connection")
        conn.Open strconn
    end sub

    sub CloseConn()
        conn.Close
        set conn = nothing
    end sub
    '--数据库连接--'

    dim strOptrCur
    dim strOptrAdminLogin
    dim strOptrAdminLogout
    dim strOptrResetDatabase
    dim strOptrImportIPTable
    dim strOptrImportMACTable
    dim strOptrImportChannelData
    dim strOptrAddVisitRule
    dim strOptrDeleteVisitRule
    dim strOptrModifyVisitRulePage
    dim strOptrModifyVisitRuleDoIt
    dim strOptrClearVisitRule
    dim strOptrDeleteVisitRecord
    dim strOptrClearVisitRecord
    dim strOptrExportVisitRecord
    dim strOptrVisitRecordCond
    dim strOptrTablePageSubmit
    dim strOptrExportMACPermTable
    dim strOptrAutoPermitMAC

    strOptrCur                = Request("optr")
    strOptrAdminLogin         = "1"
    strOptrAdminLogout        = "2"
    strOptrResetDatabase      = "3"
    strOptrImportIPTable      = "4"
    strOptrImportMACTable     = "5"
    strOptrImportChannelData  = "6"
    strOptrAddVisitRule       = "7"
    strOptrDeleteVisitRule    = "8"
    strOptrModifyVisitRulePage= "9"
    strOptrModifyVisitRuleDoIt= "a"
    strOptrClearVisitRule     = "b"
    strOptrDeleteVisitRecord  = "c"
    strOptrClearVisitRecord   = "d"
    strOptrExportVisitRecord  = "e"
    strOptrVisitRecordCond    = "f"
    strOptrTablePageSubmit    = "g"
    strOptrExportMACPermTable = "h"
    strOptrAutoPermitMAC      = "i"
%>

