<%
	dim conn
	dim connstr
	dim db
	'�������ݿ�����·��
	db="database/DB.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.Open connstr
	'�������ݿ�����
	Response.Expires=-1
	Response.ExpiresAbsolute=Now()-1
	Response.cachecontrol="no-cache"
	set rs=server.createobject("ADODB.recordset")
%>