<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="cnn.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">


<head>
<title>������ά�����ն˼��</title>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
    <meta http-equiv="refresh" content="5" />
    <!-- ˢ��ҳ���ʱ���� -->
</head>
<body>

<%
response.Write "��¼ʱ��"
response.Write " "
response.Write "����ʱ��"
response.Write " "
response.Write "��ά�ն�����"
response.Write " "
response.Write "��������"
response.Write " "
response.Write "�ն�IP"
response.Write " "
response.Write "<br>"
SQL="select * from dataTable order by ��¼ʱ��"
rs.open SQL,conn,1,1
response.Write Cstr(rs.recordcount)
if rs.recordcount>0 then
	'��ֵ
	for i=0 to rs.recordcount-1
		response.Write rs("��¼ʱ��")
		response.Write " "
		response.Write rs("����ʱ��")
		response.Write " "
		response.Write rs("��ά�ն�����")
		response.Write " "
		response.Write rs("��������")
		response.Write " "
		response.Write rs("�ն�IP")
		response.Write " "
		response.Write "<br>"
		rs.moveNext()
	next
end if
%>

</body>
</html>
