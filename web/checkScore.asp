<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/teach.asp" -->
<%
Dim checkScore
Dim checkScore_cmd
Dim checkScore_numRows

Set checkScore_cmd = Server.CreateObject ("ADODB.Command")
checkScore_cmd.ActiveConnection = MM_teach_STRING
checkScore_cmd.CommandText = "SELECT * FROM dbo.Score" 
checkScore_cmd.Prepared = true

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
  <div align="center">
    <h3>查询学生成绩</h3>
    <p>在列表中寻找需要查询的学生学号</p>
  </div>
</div>
<form id="form1" name="form1" method="post" action="/outScore.asp">
  <table width="400" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr>
      <td width="205" align="right"><label>
      <select name="studentID" id="studentID">
        <%
While (NOT checkScore.EOF)
%>
        <option value="<%=(checkScore.Fields.Item("studentID").Value)%>" <%If (Not isNull((checkScore.Fields.Item("studentID").Value))) Then If (CStr(checkScore.Fields.Item("studentID").Value) = CStr((checkScore.Fields.Item("studentID").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(checkScore.Fields.Item("studentID").Value)%></option>
        <%
  checkScore.MoveNext()
Wend
If (checkScore.CursorType > 0) Then
  checkScore.MoveFirst
Else
  checkScore.Requery
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
checkScore.Close()
Set checkScore = Nothing
%>
