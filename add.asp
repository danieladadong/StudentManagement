<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>添加学生信息</title>
</head>
<form action="add.asp" method="post" style="width:150px; position:absolute;top:20px; left:40%;">
	<fieldset style="width:150px;">
	学号：<input type="text" name="sid" /><br />
    姓名：<input type="text" name="sname" /><br />
    性别：<input type="text" name="ssex" /><br />
    班级：<input type="text" name="sclass" /><br />
    专业：<input type="text" name="smajor" /><br />
    学校：<input type="text" name="sschool" /><br />
    教师姓名：<input type="text" name="steacher" /><br />
    <input type="submit" name="add" value="添加" />
    <input type="reset" style="margin-left:49%;" value="重置" />
    </fieldset>
</form>
<%
	dim ra
	sid=request.Form("sid")
	sname=request.Form("sname")
	ssex=request.Form("ssex")
	sclass=request.Form("sclass")
	smajor=request.Form("smajor")
	sschool=request.Form("sschool")
	steacher=request.Form("steacher")
	add=request.Form("add")
	if strcomp(add,"添加")=0 then
		sql = "insert into [student] (sid,sname,ssex,sclass,smajor,sschool,steacher) values('"&sid&"','"&sname&"','"&sclass&"','"&smajor&"','"&sschool&"','"&steacher&"','"&steacher&"')"
		conn.execute sql,ra
		if ra<>0 then
			response.Redirect("index.asp")
		else
			response.Write("<script>alert('添加失败')</script>")
			response.Redirect("add.asp")
		end if
	end if
%>
<body>
</body>
</html>
