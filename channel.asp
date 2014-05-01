<!-- #include file ="common.asp" -->
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

    'open connection
    OpenConn()

    'get ip location & permission
    strIPLocation   = GetLocationByIP(strIPAddress)
    iVisitPermitted = GetPermissionByIPMAC(strIPAddress, strMACAddress)

    if iVisitPermitted = -1 then
        If strIPLocation = "China" then
            iVisitPermitted = 0
        else
            iVisitPermitted = 1
        end if
    end if

    if bSaveVisitRecord then
        WriteVisitRecord strIPAddress, strMACAddress, iVisitPermitted, strIPLocation
    end if

    'close connection
    CloseConn()

    if iVisitPermitted = 1 then
        select case nWriteOutType
        case 1
            WriteOutBinaryFile(strChannelXmlPath)
        case 2
            Response.Redirect(strChannelXmlPath)
        end select
    end if
%>

<%
    function GetLocationByIP(ip)
        dim a, num_ip, rs, sql
        a = split(ip, ".")
        num_ip = ((cdbl(a(0)) * 256 + cdbl(a(1))) * 256 + cdbl(a(2))) * 256 + cdbl(a(3))

        set rs = Server.CreateObject("ADODB.recordset")
        sql = "SELECT LOCATION FROM IPLocationTable WHERE NUM_IP_START <= " & num_ip
        sql = sql & " AND NUM_IP_END >= " & num_ip
        rs.Open sql, conn
    
        if not rs.EOF then
            GetLocationByIP = rs("LOCATION")
        else
            GetLocationByIP = "unknown"
        end If

        rs.Close()
        set rs = nothing
    end function

    function GetPermissionByIPMAC(ip, mac)
        dim rs, sql
        set rs = Server.CreateObject("ADODB.recordset")
        sql = "SELECT VisitPermission FROM VisitRuleTable WHERE IP = '" & ip & "' AND MAC = '" & mac & "'"
        sql = sql & " OR mac = '" & mac & "' AND IP = '*' "
        sql = sql & " OR ip = '" & ip & "' AND MAC = '*' "
        sql = sql & " OR ip = '*' AND MAC = '*'"
        sql = sql & "ORDER BY MAC, IP"
        rs.Open sql, conn, 1

        if rs.EOF then
            GetPermissionByIPMAC = -1
        else
            rs.MoveLast()
            GetPermissionByIPMAC = rs("VisitPermission")
        end if

        rs.Close()
        set rs = nothing
    end function
  
    sub WriteVisitRecord(ip, mac, perm, loc)
        dim rs, sql
        set rs = Server.CreateObject("ADODB.recordset")
        sql = "SELECT * FROM VisitRecordTable WHERE IP = '" & ip & "'"
        sql = sql & " AND MAC = '" & mac & "'"
        rs.Open sql, conn, 1, 3
        if rs.EOF then
            rs.AddNew()
            rs("IP" ) = ip
            rs("MAC") = mac
        end if

        rs("VisitCounter"   ) = rs("VisitCounter") + 1
        rs("VisitLastTime"  ) = now()
        rs("VisitPermission") = perm
        rs("VisitLocation"  ) = loc
        on error resume next
        rs.Update()
        rs.Close()
        set rs = nothing
    end sub

    sub WriteOutBinaryFile(url)
        dim ado
        set ado  = Server.CreateObject("Adodb.Stream")
        ado.Type = 1
        ado.Open()
        ado.LoadFromFile(Server.MapPath(url))

        Response.AddHeader "Content-Length", ado.size
        Response.ContentType = "application/octet-stream"
        Response.BinaryWrite ado.Read
        Response.Flush

        ado.Close()
        set ado = nothing
    end sub

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
%>

