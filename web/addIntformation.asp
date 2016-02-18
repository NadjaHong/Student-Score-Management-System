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
    MM_editCmd.CommandText = "INSERT INTO dbo.Student (studentID, studentName, sex, birthdate, hometown, school, major, department, native) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("studentID"), Request.Form("studentID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 10, Request.Form("studentName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 10, Request.Form("sex")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 135, 1, -1, MM_IIF(Request.Form("birthdate"), Request.Form("birthdate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 20, Request.Form("hometown")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 20, Request.Form("school")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 202, 1, 20, Request.Form("major")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 202, 1, 20, Request.Form("department")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 135, 1, -1, MM_IIF(Request.Form("native"), Request.Form("native"), null)) ' adDBTimeStamp
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
.STYLE2 {color: #000000}
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
  <div align="center">
    <h3>录入学生信息</h3>
    <p>在文本框内输入相应的信息</p>
  </div>
</div>
<form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <table width="400" height="300" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>学号:</strong></div></td>
      <td><input name="studentID" type="text" id="studentID" value="" size="25" />        
        <span class="STYLE1">* </span>11位数字 </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>姓名:</strong></div></td>
      <td><input name="studentName" type="text" id="studentName" value="" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>性别:</strong></div></td>
      <td><input name="sex" type="text" id="sex" value="" size="25" />
      <span class="STYLE1">* </span> 男/女</td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>出生日期:</strong></div></td>
      <td><input name="birthdate" type="text" id="birthdate" value="" size="25" />
      <span class="STYLE1">* </span> 如1900/1/1</td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>籍贯:</strong></div></td>
      <td><input name="hometown" type="text" id="hometown" value="" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>学校:</strong></div></td>
      <td><input name="school" type="text" id="school" value="" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>专业:</strong></div></td>
      <td><input name="major" type="text" id="major" value="" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>所在系:</strong></div></td>
      <td><input name="department" type="text" id="department" value="" size="25" />
      <span class="STYLE1">* </span> </td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right"><div align="left"><strong>入学时间:</strong></div></td>
      <td><input name="native" type="text" id="native" value="" size="25" />
      <span class="STYLE1">* </span> 如1900/1/1</td>
    </tr>
    <tr valign="baseline">
      <td colspan="2" align="center" nowrap="nowrap"><strong>
        <input type="submit" onclick="MM_validateForm('studentID','','RisNum','studentName','','R','sex','','R','birthdate','','R','hometown','','R','school','','R','major','','R','department','','R','native','','R');return document.MM_returnValue" value="录入" />
        <label>
        <input type="reset" name="button" id="button" value="重置" />
          </label>
            </strong></td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1" />
</form>
<p>&nbsp;</p>
</body>
</html>
