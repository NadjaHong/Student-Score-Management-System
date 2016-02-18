<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/teach.asp" -->
<%
Dim checkScore__MMColParam
checkScore__MMColParam = "1"
If (Request.Form("studentID") <> "") Then 
  checkScore__MMColParam = Request.Form("studentID")
End If
%>
<%
Dim checkScore
Dim checkScore_cmd
Dim checkScore_numRows

Set checkScore_cmd = Server.CreateObject ("ADODB.Command")
checkScore_cmd.ActiveConnection = MM_teach_STRING
checkScore_cmd.CommandText = "SELECT * FROM dbo.Score WHERE studentID = ?" 
checkScore_cmd.Prepared = true
checkScore_cmd.Parameters.Append checkScore_cmd.CreateParameter("param1", 5, 1, -1, checkScore__MMColParam) ' adDouble

Set checkScore = checkScore_cmd.Execute
checkScore_numRows = 0
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
  <h3 align="center">查询学生成绩</h3>
</div>
<form id="form1" name="form1" method="post" action="">
  <table width="400" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr>
      <td width="130"><strong>学号：</strong></td>
      <td width="354"><%=(checkScore.Fields.Item("studentID").Value)%></td>
    </tr>
    <tr>
      <td><strong>姓名：</strong></td>
      <td><%=(checkScore.Fields.Item("studentName").Value)%></td>
    </tr>
    <tr>
      <td><strong>课程名称：</strong></td>
      <td><%=(checkScore.Fields.Item("courseName").Value)%></td>
    </tr>
    <tr>
      <td><strong>成绩：</strong></td>
      <td><%=(checkScore.Fields.Item("score").Value)%></td>
    </tr>
  </table>
</form>  
  <div>
    <div align="center"><strong><a href="/checkScore.asp">返回</a></strong></div>
  </div>
</body>
</html>
<%
checkScore.Close()
Set checkScore = Nothing
%>
