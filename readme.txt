文件说明：

tvbox.mdb        数据库文件
channel.asp      处理频道请求的程序
channel.xml      存放频道数据的文件
admin.asp        管理页面程序
submit.asp       数据处理程序
help.html        使用说明
iptable.txt      ip 地址库  - 建议改为 *.asp 这样的名字更安全
mactable.txt     mac 授权库 - 建议改为 *.asp 这样的名字更安全


配置说明：

所有配置项都在 common.asp 中

strChannelXmlPath = "channel.xml"   // 真实频道数据文件名
strIPLocTableTxt  = "iptable.txt"   // ip 地址库文件，文本形式，可导入数据库来更新 IP 地址表
strMACTableTxt    = "mac.txt"       // mac 授权库文件，文本形式，可导入数据库来更新 MAC 授权表
strAdminLoginPwd  = "www.tvbox.com" // 管理员登陆密码
strAdminPageName  = "admin.asp"     // 管理页面名，如果修改了 admin.asp 的名称，需求修改这个
strLoginPageName  = "login.asp"     // 管理页面名，如果修改了 login.asp 的名称，需求修改这个
bSaveVisitRecord  = true            // 是否保存访问记录
nDataBaseType     = 1               // 数据库类型，1-access, 2/3-sqlserver
nWriteOutType     = 1               // 频道数据写出方式 1-write binary 方式，2-重定向方式
nTablePageSize    = 10              // 每页显示的记录个数
strChinaCode      = "China"         // 中国的国家代码，需要根据实际使用的 ip 库进行配置

数据库初始化
strAccessDBFile  = "tvbox.mdb"         // access 数据库文件名
strAccessDBPWD   = "www.tvbox.com"     // access 数据库密码，默认 www.tvbox.com

strSQLServerHost = "WIN-GDONNGTOTGF\LOOLBOX_4_29" // sqlserver host
strSQLServerDBN  = "looltv_content"               // sqlserver database
strSQLServerUSER = "sa"                           // sqlserver user
strSQLServerPWD  = "Loolbox2014"                  // sqlserver password

第一次使用需要先执行 initdb.asp 进行数据库初始化，以后就不要执行了，否则会导致数据清空。


iptool.exe 使用说明：
用于进行 ip 库转换为 tvlinkperm 系统能导入数据库的 iptable.txt 文件
在同一级目录下放一个名为 ip.txt 的输入文件，双击执行 iptool.exe 即可生成 ip.out.txt
将 ip.out.txt 改名为 iptable.txt 即可。

