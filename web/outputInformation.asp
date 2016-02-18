<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/teach.asp" -->
<%
Dim checkInformation__MMColParam
checkInformation__MMColParam = "1"
If (Request.Form("studentID") <> "") Then 
  checkInformation__MMColParam = Request.Form("studentID")
End If
%>
<%
Dim checkInformation
Dim checkInformation_cmd
Dim checkInformation_numRows

Set checkInformation_cmd = Server.CreateObject ("ADODB.Command")
checkInformation_cmd.ActiveConnection = MM_teach_STRING
checkInformation_cmd.CommandText = "SELECT * FROM dbo.Student WHERE studentID = ?" 
checkInformation_cmd.Prepared = true
checkInformation_cmd.Parameters.Append checkInformation_cmd.CreateParameter("param1", 5, 1, -1, checkInformation__MMColParam) ' adDouble

Set checkInformation = checkInformation_cmd.Execute
checkInformation_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
<style type="text/css">
<!--
body {
	background-image: url(/2.jpg);
}
-->
</style></head>

<body>
<div>
  <h3 align="center">查询学生信息</h3>
</div>
<form action="" method="get">
<table width="400" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
  <tr>
    <td><strong>学号：</strong></td>
    <td><%=(checkInformation.Fields.Item("studentID").Value)%></td>
  </tr>
  <tr>
    <td><strong>姓名：</strong></td>
    <td><%=(checkInformation.Fields.Item("studentName").Value)%></td>
  </tr>
  <tr>
    <td><strong>性别：</strong></td>
    <td><%=(checkInformation.Fields.Item("sex").Value)%></td>
  </tr>
  <tr>
    <td><strong>出生日期：</strong></td>
    <td><%=(checkInformation.Fields.Item("birthdate").Value)%></td>
  </tr>
  <tr>
    <td><strong>籍贯：</strong></td>
    <td><%=(checkInformation.Fields.Item("hometown").Value)%></td>
  </tr>
  <tr>
    <td><strong>学校：</strong></td>
    <td><%=(checkInformation.Fields.Item("school").Value)%></td>
  </tr>
  <tr>
    <td><strong>专业：</strong></td>
    <td><%=(checkInformation.Fields.Item("major").Value)%></td>
  </tr>
  <tr>
    <td><strong>所在系：</strong></td>
    <td><%=(checkInformation.Fields.Item("department").Value)%></td>
  </tr>
  <tr>
    <td><strong>入学时间：</strong></td>
    <td><%=(checkInformation.Fields.Item("native").Value)%></td>
  </tr>
</table>
</form>
  <div>
    <div align="center"><strong><a href="/checkIntformation.asp">返回</a></strong></div>
  </div>
</body>
</html>
<%
checkInformation.Close()
Set checkInformation = Nothing
%>
