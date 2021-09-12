<%
	dim conn
	dim connstr
	dim db
	'设置数据库连接路径
	db="database/DB.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.Open connstr
	'创建数据库连接
	Response.Expires=-1
	Response.ExpiresAbsolute=Now()-1
	Response.cachecontrol="no-cache"
	set rs=server.createobject("ADODB.recordset")
%>