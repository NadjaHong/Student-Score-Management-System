<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/teach.asp" -->
<%
Dim checkInformation
Dim checkInformation_cmd
Dim checkInformation_numRows

Set checkInformation_cmd = Server.CreateObject ("ADODB.Command")
checkInformation_cmd.ActiveConnection = MM_teach_STRING
checkInformation_cmd.CommandText = "SELECT * FROM dbo.Student" 
checkInformation_cmd.Prepared = true

Set checkInformation = checkInformation_cmd.Execute
checkInformation_numRows = 0
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
  <p align="center">在列表中寻找需要查询的学生学号</p>
</div>
<form id="form1" name="form1" method="post" action="/outputInformation.asp">
  <table width="400" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr>
      <td width="205" align="right"><label>
      <select name="studentID" id="studentID">
        <%
While (NOT checkInformation.EOF)
%><option value="<%=(checkInformation.Fields.Item("studentID").Value)%>" <%If (Not isNull((checkInformation.Fields.Item("studentID").Value))) Then If (CStr(checkInformation.Fields.Item("studentID").Value) = CStr((checkInformation.Fields.Item("studentID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(checkInformation.Fields.Item("studentID").Value)%></option>
        <%
  checkInformation.MoveNext()
Wend
If (checkInformation.CursorType > 0) Then
  checkInformation.MoveFirst
Else
  checkInformation.Requery
End If
%>
        </select>
      </label></td>
      <td width="179"><label>
        <input type="submit" name="button" id="button" value="查询" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
checkInformation.Close()
Set checkInformation = Nothing
%>
