�ļ�˵����

tvbox.mdb        ���ݿ��ļ�
channel.asp      ����Ƶ������ĳ���
channel.xml      ���Ƶ�����ݵ��ļ�
admin.asp        ����ҳ�����
submit.asp       ���ݴ������
help.html        ʹ��˵��
iptable.txt      ip ��ַ��  - �����Ϊ *.asp ���������ָ���ȫ
mactable.txt     mac ��Ȩ�� - �����Ϊ *.asp ���������ָ���ȫ


����˵����

����������� common.asp ��

strChannelXmlPath = "channel.xml"   // ��ʵƵ�������ļ���
strIPLocTableTxt  = "iptable.txt"   // ip ��ַ���ļ����ı���ʽ���ɵ������ݿ������� IP ��ַ��
strMACTableTxt    = "mac.txt"       // mac ��Ȩ���ļ����ı���ʽ���ɵ������ݿ������� MAC ��Ȩ��
strAdminLoginPwd  = "www.tvbox.com" // ����Ա��½����
strAdminPageName  = "admin.asp"     // ����ҳ����������޸��� admin.asp �����ƣ������޸����
strLoginPageName  = "login.asp"     // ����ҳ����������޸��� login.asp �����ƣ������޸����
bSaveVisitRecord  = true            // �Ƿ񱣴���ʼ�¼
nDataBaseType     = 1               // ���ݿ����ͣ�1-access, 2/3-sqlserver
nWriteOutType     = 1               // Ƶ������д����ʽ 1-write binary ��ʽ��2-�ض���ʽ
nTablePageSize    = 10              // ÿҳ��ʾ�ļ�¼����
strChinaCode      = "China"         // �й��Ĺ��Ҵ��룬��Ҫ����ʵ��ʹ�õ� ip ���������

���ݿ��ʼ��
strAccessDBFile  = "tvbox.mdb"         // access ���ݿ��ļ���
strAccessDBPWD   = "www.tvbox.com"     // access ���ݿ����룬Ĭ�� www.tvbox.com

strSQLServerHost = "WIN-GDONNGTOTGF\LOOLBOX_4_29" // sqlserver host
strSQLServerDBN  = "looltv_content"               // sqlserver database
strSQLServerUSER = "sa"                           // sqlserver user
strSQLServerPWD  = "Loolbox2014"                  // sqlserver password

��һ��ʹ����Ҫ��ִ�� initdb.asp �������ݿ��ʼ�����Ժ�Ͳ�Ҫִ���ˣ�����ᵼ��������ա�


iptool.exe ʹ��˵����
���ڽ��� ip ��ת��Ϊ tvlinkperm ϵͳ�ܵ������ݿ�� iptable.txt �ļ�
��ͬһ��Ŀ¼�·�һ����Ϊ ip.txt �������ļ���˫��ִ�� iptool.exe �������� ip.out.txt
�� ip.out.txt ����Ϊ iptable.txt ���ɡ�

