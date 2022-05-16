<%
' FileName="Connection_ado_conn_string.htm"
' Type="ADO" 
' DesigntimeType="ADO"
' HTTP="false"
' Catalog=""
' Schema=""
dim conn,rs,sql 
'on error resume next 
'dbpath为你自己设置的数据库路径
dbpath="E:/student/student.mdb" 
set conn=Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; data source="&dbpath
'创建记录对象 
set rs=server.createobject("adodb.recordset") 
%>
