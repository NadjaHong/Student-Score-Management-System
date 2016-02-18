<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/teach.asp" -->
<%
Dim updateface
Dim updateface_cmd
Dim updateface_numRows

Set updateface_cmd = Server.CreateObject ("ADODB.Command")
updateface_cmd.ActiveConnection = MM_teach_STRING
updateface_cmd.CommandText = "SELECT * FROM dbo.Score" 
updateface_cmd.Prepared = true

Set updateface = updateface_cmd.Execute
updateface_numRows = 0
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
  <h3 align="center">修改学生成绩</h3>
  <p align="center">在文本框内输入需要修改的学生学号</p>
</div>
<form id="form1" name="form1" method="post" action="/updateScore.asp">
  <table width="400" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr>
      <td align="center"><label>
      <select name="studentID" id="studentID">
        <%
While (NOT updateface.EOF)
%>
        <option value="<%=(updateface.Fields.Item("studentID").Value)%>"><%=(updateface.Fields.Item("studentID").Value)%></option>
        <%
  updateface.MoveNext()
Wend
If (updateface.CursorType > 0) Then
  updateface.MoveFirst
Else
  updateface.Requery
End If
%>
      </select>
      </label></td>
      <td><label>
        <input type="submit" name="button" id="button" value="修改" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
updateface.Close()
Set updateface = Nothing
%>
