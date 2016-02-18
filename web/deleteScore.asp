<%@LANGUAGE="VBSCRIPT"%>
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
%><%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form2" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_teach_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.Score WHERE studentID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "/deleteface.asp"
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
Dim deleteScore
Dim deleteScore_cmd
Dim deleteScore_numRows

Set deleteScore_cmd = Server.CreateObject ("ADODB.Command")
deleteScore_cmd.ActiveConnection = MM_teach_STRING
deleteScore_cmd.CommandText = "SELECT * FROM dbo.Score" 
deleteScore_cmd.Prepared = true

Set deleteScore = deleteScore_cmd.Execute
deleteScore_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
<!--
body {
	background-image: url(/2.jpg);
}
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
  <div align="center">
    <h3>删除学生成绩</h3>
  </div>
</div>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form2" id="form2">
  <table width="400" height="150" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr valign="baseline">
      <td width="100" align="right" nowrap="nowrap"><div align="left">
        <h3><strong>学号:</strong></h3>
      </div></td>
      <td width="300"><%=(deleteScore.Fields.Item("studentID").Value)%> </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="left">
        <h3><strong>姓名:</strong></h3>
      </div></td>
      <td><%=(deleteScore.Fields.Item("studentName").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="left">
        <h3><strong>课程名称:</strong></h3>
      </div></td>
      <td><%=(deleteScore.Fields.Item("courseName").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="left">
        <h3><strong>成绩:</strong></h3>
      </div></td>
      <td><%=(deleteScore.Fields.Item("score").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td colspan="2" align="center" nowrap="nowrap"><input type="submit" onclick="MM_popupMsg('删除成功')" value="确定删除" />
      <label>  
      <a href="/deleteface.asp"><input type="button" name="button" id="button" value="取消" /></a>      </label></td>
    </tr>
  </table>
  

  <input type="hidden" name="MM_delete" value="form2" />
  <input type="hidden" name="MM_recordId" value="<%= deleteScore.Fields.Item("studentID").Value %>" />
</form>
<p>&nbsp;</p>
</body>
</html>
<%
deleteScore.Close()
Set deleteScore = Nothing
%>
