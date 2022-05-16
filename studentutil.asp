<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/conn.asp" -->
<%
	Session.CodePage=65001 
	Response.Charset="UTF-8"
	sid = request.Form("sid")
	flag = request.Form("flag")
	if strcomp(flag,"删除")=0 then
		sql = "delete from student where [sid]="&sid
		conn.execute sql
		response.Redirect("index.asp")
	else
		dim ra
		sname = request.Form("sname")
		ssex = request.Form("ssex")
		sclass = request.Form("sclass")
		smajor = request.Form("smajor")
		sschool = request.Form("sschool")
		steacher = request.Form("steacher")
		sql = "update student set[sname]='"&sname&"',[ssex]='"&ssex&"',[sclass]='"&sclass&"',[smajor]='"&smajor&"',[sschool]='"&sschool&"',[steacher]='"&steacher&"' where [sid]="&sid
		conn.execute sql,ra
		if ra<>0 then
			response.Redirect("index.asp")
		end if
	end if

%>