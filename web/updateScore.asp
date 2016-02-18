<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/teach.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_teach_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Score SET studentID = ?, studentName = ?, courseName = ?, score = ? WHERE studentID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("studentID"), Request.Form("studentID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 10, Request.Form("studentName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 20, Request.Form("courseName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("score"), Request.Form("score"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "/updateface.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim updateScore__MMColParam
updateScore__MMColParam = "1"
If (Request.Form("studentID") <> "") Then 
  updateScore__MMColParam = Request.Form("studentID")
End If
%>
<%
Dim updateScore
Dim updateScore_cmd
Dim updateScore_numRows

Set updateScore_cmd = Server.CreateObject ("ADODB.Command")
updateScore_cmd.ActiveConnection = MM_teach_STRING
updateScore_cmd.CommandText = "SELECT * FROM dbo.Score WHERE studentID = ?" 
updateScore_cmd.Prepared = true
updateScore_cmd.Parameters.Append updateScore_cmd.CreateParameter("param1", 5, 1, -1, updateScore__MMColParam) ' adDouble

Set updateScore = updateScore_cmd.Execute
updateScore_numRows = 0
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
.STYLE1 {color: #FF0000}
-->
</style>
<script type="text/javascript">
<!--
function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
//-->
</script>
</head>

<body>
<div>
  <h3 align="center">修改学生成绩</h3>
  <p align="center">在文本框内输入相应的信息</p>
</div>
<form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <table width="400" height="150" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>学号:</strong></div></td>
      <td><input name="studentID" type="text" value="<%=(updateScore.Fields.Item("studentID").Value)%>" size="25" />
      <span class="STYLE1">* </span> 11位数字</td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>姓名:</strong></div></td>
      <td><input name="studentName" type="text" value="<%=(updateScore.Fields.Item("studentName").Value)%>" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>课程名称:</strong></div></td>
      <td><input name="courseName" type="text" value="<%=(updateScore.Fields.Item("courseName").Value)%>" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>成绩:</strong></div></td>
      <td><input name="score" type="text" value="<%=(updateScore.Fields.Item("score").Value)%>" size="25" />      </td>
    </tr>
    <tr valign="baseline">
      <td colspan="2" align="center" nowrap="nowrap"><input name="提交" type="submit" onclick="MM_popupMsg('更新成功')" value="更新" />
        <label>
        <input type="reset" name="button" id="button" value="重置" />
      </label>      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1" />
  <input type="hidden" name="MM_recordId" value="<%= updateScore.Fields.Item("studentID").Value %>" />
</form>
</body>
</html>
<%
updateScore.Close()
Set updateScore = Nothing
%>
