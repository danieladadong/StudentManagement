<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="Connections/conn.asp" -->
<%
Session.CodePage=65001 
Response.Charset="UTF-8"
username = request.Form("username")
password = request.Form("password")
set rsc=server.createobject("adodb.recordset")
sqlc="select * from [suser] where [username]='"&username&"' and [password]='"&password&"'"
rsc.open sqlc,conn,1,1
if rsc.eof then
	Response.Write("用户名或密码错误！")
	Response.Write("<br><a href='index.html'>index.html</a><br />")
	Response.End()
	else
	session("username")=rsc("username")
	response.Redirect("index.asp")
end if
	
%>
