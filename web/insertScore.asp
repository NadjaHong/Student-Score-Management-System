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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_teach_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Score (studentID, studentName, courseName, score) VALUES (?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("studentID"), Request.Form("studentID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 10, Request.Form("studentName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 20, Request.Form("courseName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("score"), Request.Form("score"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
  End If
End If
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
function MM_validateForm() { //v4.0
  if (document.getElementById){
    var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
    for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=document.getElementById(args[i]);
      if (val) { nm=val.name; if ((val=val.value)!="") {
        if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
          if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
        } else if (test!='R') { num = parseFloat(val);
          if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
          if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
            min=test.substring(8,p); max=test.substring(p+1);
            if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
      } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
    } if (errors) alert('The following error(s) occurred:\n'+errors);
    document.MM_returnValue = (errors == '');
} }
//-->
</script>
</head>

<body>
<div>
  <h3 align="center">添加学生成绩</h3>
  <p align="center">在文本框内输入相应的信息</p>
</div>
<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
<table width="400" height="150" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>学号:</strong></div></td>
      <td><input name="studentID" type="text" id="studentID" value="" size="25" />
        <span class="STYLE1">*</span> 11位数字</td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>姓名:</strong></div></td>
      <td><input name="studentName" type="text" id="studentName" value="" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>课程名称:</strong></div></td>
      <td><input name="courseName" type="text" id="courseName" value="" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>成绩:</strong></div></td>
      <td><input type="text" name="score" value="" size="25" /></td>
    </tr>
    <tr valign="baseline">
      <td colspan="2" align="center" nowrap="nowrap"><strong>
        <input name="提交" type="submit" onclick="MM_validateForm('studentID','','RisNum','studentName','','R','courseName','','R');return document.MM_returnValue" value="添加" />
        <label>
        <input type="reset" name="button" id="button" value="重置" />
          </label>
            </strong></td>
    </tr>
  </table>

  
  <input type="hidden" name="MM_insert" value="form1" />
</form>
</body>
</html>
