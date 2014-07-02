<%@ Import Namespace="System.data" %>
<%@ import namespace="system.data.oledb" %>
<%@ import namespace="mysql.data.mysqlclient" %>
<%@ page language="vb" debug="true" %>

<%
    dim strChannelXmlPath
    dim bSaveVisitRecord
    dim nWriteOutType
    dim strChinaCode

    strChannelXmlPath = "channel.xml"
    bSaveVisitRecord  = true
    nWriteOutType     = 1
    strChinaCode      = "China"

    dim nDataBaseType
    dim strAccessDBFile
    dim strAccessDBPWD
    dim strSQLServerHost
    dim strSQLServerUSER
    dim strSQLServerPWD
    dim strSQLServerDBN

    'database type
    nDataBaseType    = 1

    'access
    strAccessDBFile  = "tvbox.mdb"
    strAccessDBPWD   = "www.tvbox.com"

    'sqlserver
    strSQLServerHost = "WIN-GDONNGTOTGF\LOOLBOX_4_29"
    strSQLServerUSER = "sa"
    strSQLServerPWD  = "Loolbox2014"
    strSQLServerDBN  = "looltv_content"

    dim dbconn, strconn
    select case nDataBaseType
    case 1
        strconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(strAccessDBFile) & ";"
        strconn = strconn & "Persist Security Info=false;Jet OLEDB:Database Password=" & strAccessDBPWD
    case 2
        strconn = "provider=sqloledb;"
        strconn = strconn & "data source=" & strSQLServerHost & ";"
        strconn = strconn & "initial catalog=" & strSQLServerDBN & ";"
        strconn = strconn & "user id=" & strSQLServerUSER & ";"
        strconn = strconn & "password=" & strSQLServerPWD
    end select
%>

<script language="vb" runat="server">
    function IsValidMACAddress(mac)
        if len(mac) <> 17 then
            IsValidMACAddress = false
            exit function
        end if
        if instr(mac, "'") <> 0 then
            IsValidMACAddress = false
            exit function
        end if
        IsValidMACAddress = true
    end function

    function GetLocationByIPAddr(conn, ip)
        dim a, num_ip, sql, cmd, dr
        a = split(ip, ".")
        num_ip = ((cdbl(a(0)) * 256 + cdbl(a(1))) * 256 + cdbl(a(2))) * 256 + cdbl(a(3))
        sql = "SELECT LOCATION FROM IPLocationTable WHERE NUM_IP_START <= " & num_ip
        sql = sql & " AND NUM_IP_END >= " & num_ip
        cmd = new OleDbCommand(sql, conn)
        dr  = cmd.ExecuteReader()
        if dr.Read() then
            GetLocationByIPAddr = dr("LOCATION")
        else
            GetLocationByIPAddr = "unknown"
        end if
        dr.close()
    end function

    function GetPermissionByRule(conn, ip, mac)
        dim sql, cmd, dr
        sql = "SELECT VisitPermission FROM VisitRuleTable WHERE IP = '" & ip & "' AND MAC = '" & mac & "'"
        sql = sql & " OR mac = '" & mac & "' AND IP = '*' "
        sql = sql & " OR ip = '" & ip & "' AND MAC = '*' "
        sql = sql & " OR ip = '*' AND MAC = '*'"
        sql = sql & "ORDER BY MAC DESC, IP DESC"
        cmd = new OleDbCommand(sql, conn)
        dr  = cmd.ExecuteReader()

        if dr.Read() then
            GetPermissionByRule = dr("VisitPermission")
        else
            GetPermissionByRule = -1
        end if
        dr.close()
    end function

    sub WriteVisitRecord(conn, ip, mac, perm, loc)
        dim sql, cmd, da, dt, row, cmdb
        sql = "SELECT * FROM VisitRecordTable WHERE IP = '" & ip & "'"
        sql = sql & " AND MAC = '" & mac & "'"
        cmd = new OleDbCommand(sql, conn)
        da  = new OleDbDataAdapter(cmd)
        cmdb= new OleDbCommandBuilder(da)
        dt  = new DataTable()

        da.Fill(dt)
        if dt.Rows.Count = 0 then
            row = dt.NewRow()
            row("IP" ) = ip
            row("MAC") = mac
            row("VisitCounter"   ) = 1
            row("VisitLastTime"  ) = now()
            row("VisitPermission") = perm
            row("VisitLocation"  ) = loc
            dt.Rows.Add(row)
        else
            dt.Rows(0)("VisitCounter"   ) = dt.Rows(0)("VisitCounter") + 1
            dt.Rows(0)("VisitLastTime"  ) = now()
            dt.Rows(0)("VisitPermission") = perm
            dt.Rows(0)("VisitLocation"  ) = loc
        end if
        try
            da.Update(dt)
        catch e as exception
        end try
    end sub

    function GetPermissionByMAC(strMACAddress)
        dim conn, cnstr, sql, cmd, dr
        cnstr = "server=localhost;port=8306;database=forcecms_4_0_5;user id=root;password=forcetech"
        conn  = new MySqlConnection(cnstr)
        conn.Open()

        sql = "SELECT * FROM f_device WHERE MacAddress='" & strMACAddress & "'"
        cmd = new MySqlCommand(sql, conn)
        dr  = cmd.ExecuteReader()

        if dr.Read() then
            GetPermissionByMAC = 1
        else
            GetPermissionByMAC = 0
        end if

        dr.Close()
        conn.Close()
    end function
</script>

<%
    dim strIPAddress
    dim strMACAddress
    dim strIPLocation
    dim iVisitPermitted

    strIPAddress    = "127.0.0.1"
    strMACAddress   = "00:00:00:00:00:00"
    strIPLocation   = "unknown"
    iVisitPermitted = -1

    'get ip address
    strIPAddress = Request.ServerVariables("http_x_forwarded_for")
    if strIPAddress = "" then strIPAddress = Request.ServerVariables("Remote_Addr")

    'get mac address
    strMACAddress = lcase(Request.QueryString("mac"))
    if not IsValidMACAddress(strMACAddress) then strMACAddress = "00:00:00:00:00:00"

    'open conn
    dbconn = new OleDBConnection(strconn)
    dbconn.Open()

    'get ip location & permission
    strIPLocation   = GetLocationByIPAddr(dbconn, strIPAddress)
    iVisitPermitted = GetPermissionByRule(dbconn, strIPAddress, strMACAddress)

    if iVisitPermitted = -1 then
        If strIPLocation = strChinaCode then
            iVisitPermitted = 0
        else
            iVisitPermitted = GetPermissionByMAC(strMACAddress)
        end if
    end if

    if bSaveVisitRecord then
        WriteVisitRecord(dbconn, strIPAddress, strMACAddress, iVisitPermitted, strIPLocation)
    end if

    'close conn
    dbconn.Close()

    if iVisitPermitted = 1 then
        select case nWriteOutType
        case 1
            Response.ClearContent()
            Response.AppendHeader("Content-Disposition", "attachment;filename=channel.bin")
            Response.TransmitFile(strChannelXmlPath)
        case 2
            Response.ClearContent()
            Response.AppendHeader("Content-Disposition", "attachment;filename=channel.bin")
            Response.WriteFile(strChannelXmlPath)
        case 3
            Response.ClearContent()
            Response.Redirect(strChannelXmlPath)
        end select
    end if
%>

