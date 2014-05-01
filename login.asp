<!-- #include file ="common.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0
  Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
  <meta http-equiv="Content-Language" content="zh-cn" />
  <title>管理登录</title>
</head>

<body>
<br/>
<br/>
<br/>
<br/>
<br/>
<form action="submit.asp" method="post">
  <input type="hidden" name="optr" value="<%=strOptrAdminLogin%>"/>
  <table border="1" align="center" width="35%">
    <tr><th></th><th>管理登录</th></tr>
    <tr align="center"><td>密码：</td><td><input size="32" type="password" name="pwd"/></td></tr>
    <tr align="center"><td></td><td><input type="submit" value="  登  录  "/></td></tr>
  </table>
</form>
</body>

</html>
