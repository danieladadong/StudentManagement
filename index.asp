<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/conn.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理系统</title>
<link rel="stylesheet" href="css/basic.css">
</head>
<body>
<div id="mainDiv">
	<div id="topDiv">
        <div id="tmenu">
            <ul>
              <li ><a href="index.asp">首页</a></li>
             <!-- <li ><a href="">添加用户</a></li>-->
              <li ><a href="add.asp">添加数据</a></li>
            </ul>
        </div>
  </div>
	<div id="centerDiv">
    <div id="left">
      <div id="lhead">管理系统</div>
      <form action="index.asp" method="post">
        <ul>
              <li ><a id="lo" href="login.htm">登录</a></li>
              <li ><a href=""><input style="background:none;border:none;padding-top:3px;color:#FFF;" type="submit" value="信息管理" name="util"/></a></li>
              <li ><a href=""><input style="background:none;border:none;padding-top:3px;color:#FFF;" type="submit" value="信息显示" name="util"/></a></li>
              <!--<li ><a href="admin-templates">登录管理</a></li>-->
        </ul>
      </form>
    </div>
    <div id="right"> 
        <div id="current">&nbsp;&nbsp;&nbsp;&nbsp;当前位置: <%if request.Form("util")<>"" then response.Write(request.Form("util")) end if%>
		</div>
        <div id="form">
            <%
              if strcomp(request.Form("util"),"信息管理")=0 then
			  %>
              <form action="index.asp" method="post">
                <input type="hidden" name="util" value="信息管理">
                <input type="text" name="keyword">
                <input type="submit" value="搜索"/>
              </form>
              <%
			  	keyword = request.Form("keyword")
				set rs = server.CreateObject("adodb.recordset")
				if keyword<>"" then
					sql = "select * from [student] where [sid] like '%"&keyword&"%'or [sname] like '%"&keyword&"%' or [sclass] like '%"&keyword&"%' or [smajor] like '%"&keyword&"%' or [sschool] like '%"&keyword&"%' or [steacher] like '%"&keyword&"%'"
				else
					sql = "select * from [student]"
				end if
					rs.open sql,conn,1,1
			  %>
              <table width="100%"style="text-align:center;border:1px solid; margin-left:-90px; margin-top:10px;" >
                <tr>
                  <th scope="col">学号</th>
                  <th scope="col">姓名</th>
                  <th scope="col">性别</th>
                  <th scope="col">班级</th>
                  <th scope="col">专业</th>
                  <th scope="col">学校</th>
                  <th scope="col">任课老师</th>
                  <th scope="col" colspan="2" >操作</th>
                </tr>
                <%
                  do until rs.eof
                %>
                <tr>
                <form method="post" action="studentutil.asp">
                  <td><input type="text" size="10%" style="text-align:center;" name="sid" value="<%response.Write(rs("sid"))%>"  readonly></td>
                  <td><input type="text" size="10%" style="text-align:center;" name="sname" value="<%response.Write(rs("sname"))%>"></td>
                  <td><input type="text" size="10%" style="text-align:center;" name="ssex" value="<%response.Write(rs("ssex"))%>"></td>
                  <td><input type="text" size="10%" style="text-align:center;" name="sclass" value="<%response.Write(rs("sclass"))%>"></td>
                  <td><input type="text" size="10%" style="text-align:center;" name="smajor" value="<%response.Write(rs("smajor"))%>"></td>
                  <td><input type="text" size="10%" style="text-align:center;" name="sschool" value="<%response.Write(rs("sschool"))%>"></td>
                  <td><input type="text" size="10%" style="text-align:center;" name="steacher" value="<%response.Write(rs("steacher"))%>"></td>
                  <td><input type="submit" style="width:100%;" value="删除" name="flag"></td>
                  <td><input type="submit" style="width:100%;" value="修改" name="flag"></td>
                </form>
                </tr>
                <%
                  rs.movenext
                  loop
                  rs.close
                  rs=nothing
                %>
              </table>
            <%	
			elseif strcomp(request.Form("util"),"信息显示")=0 then
			%>
				<form action="index.asp" method="post">
                	<input type="hidden" name="util" value="信息显示">
                	<input type="text" name="keyword">
                    <input type="submit" value="搜索"/>
                </form>
                <%
				keyword = request.Form("keyword")
				if keyword<>"" then
					set rs = server.CreateObject("adodb.recordset")
					sql = "select * from [student] where [sid] like '%"&keyword&"%'or [sname] like '%"&keyword&"%' or [sclass] like '%"&keyword&"%' or [smajor] like '%"&keyword&"%' or [sschool] like '%"&keyword&"%' or [steacher] like '%"&keyword&"%'"
					rs.open sql,conn,1,1
				%>
					<table width="100%" style="text-align:center; border:1px solid; margin-left:-90px; margin-top:10px;">
                    <tr>
                      <th scope="col">学号</th>
                      <th scope="col">姓名</th>
                      <th scope="col">性别</th>
                      <th scope="col">班级</th>
                      <th scope="col">专业</th>
                      <th scope="col">学校</th>
                      <th scope="col">任课老师</th>
                    </tr>
                <%
                  do until rs.eof
                %>
                <tr>
                  <td><%response.Write(rs("sid"))%></td>
                  <td><%response.Write(rs("sname"))%></td>
                  <td><%response.Write(rs("ssex"))%></td>
                  <td><%response.Write(rs("sclass"))%></td>
                  <td><%response.Write(rs("smajor"))%></td>
                  <td><%response.Write(rs("sschool"))%></td>
                  <td><%response.Write(rs("steacher"))%></td>
                </tr>
                <%
                  rs.movenext
                  loop
                  rs.close
                  rs=nothing
					end if
            end if
            %>
      </div>
	</div>
  </div>
</div>

</body>
</html>