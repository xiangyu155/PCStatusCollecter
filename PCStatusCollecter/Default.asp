<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="cnn.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">


<head>
<title>集中运维操作终端监控</title>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
    <meta http-equiv="refresh" content="5" />
    <!-- 刷新页面的时间间隔 -->
</head>
<body>

<%
response.Write "记录时间"
response.Write " "
response.Write "空闲时间"
response.Write " "
response.Write "运维终端名称"
response.Write " "
response.Write "主机名称"
response.Write " "
response.Write "终端IP"
response.Write " "
response.Write "<br>"
SQL="select * from dataTable order by 记录时间"
rs.open SQL,conn,1,1
response.Write Cstr(rs.recordcount)
if rs.recordcount>0 then
	'赋值
	for i=0 to rs.recordcount-1
		response.Write rs("记录时间")
		response.Write " "
		response.Write rs("空闲时间")
		response.Write " "
		response.Write rs("运维终端名称")
		response.Write " "
		response.Write rs("主机名称")
		response.Write " "
		response.Write rs("终端IP")
		response.Write " "
		response.Write "<br>"
		rs.moveNext()
	next
end if
%>

</body>
</html>
