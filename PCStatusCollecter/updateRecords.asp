<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="cnn.asp" -->
<%
if Cstr(request("recordTime"))<>"" and Cstr(request("spareTime"))<>"" and Cstr(request("hostInfo"))<>"" and Cstr(request("hostName"))<>"" and Cstr(request("hostIP"))<>""  then
SQL="INSERT INTO dataTable (��¼ʱ��,����ʱ��,��ά�ն�����,��������,�ն�IP) VALUES ('"&Cstr(request("recordTime"))&"', '"&Cstr(request("spareTime"))&"', '"&Cstr(request("hostInfo"))&"', '"&Cstr(request("hostName"))&"', '"&Cstr(request("hostIP"))&"');"	
	'rs.close
	rs.open SQL,conn,1,3
end if
%>

