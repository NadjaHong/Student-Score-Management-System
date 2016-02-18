<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/teach.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "interface.html"
  MM_redirectLoginFailed = "loginFalse.asp"

  MM_loginSQL = "SELECT ID, password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.ID WHERE ID = ? AND password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_teach_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 5, 1, -1, MM_valUsername) ' adDouble
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 5, 1, -1, Request.Form("password")) ' adDouble
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>学生成绩管理系统-登录</title>
<style type="text/css">
<!--
body {
	background-image: url(1.jpg);
}
.STYLE1 {color: #FFFFFF}
.STYLE3 {
	font-family: "微软雅黑"
}
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
  <h1 align="center">&nbsp;</h1>
  <p align="center">&nbsp;</p>
  <h1 align="center" class="STYLE3">学生成绩管理系统</h1>
  <p align="center">&nbsp;</p>
  <p>&nbsp;</p>
</div>
<form ACTION="<%=MM_LoginAction%>" method="POST" name="form1" id="form1" onsubmit="MM_validateForm('username','','RisNum','password','','RisNum');return document.MM_returnValue">
  <table width="400" height="150" border="0" align="center" cellspacing="0">
    <tr bgcolor="#223240">
      <td height="60" colspan="2"><div align="center">
        <h3 class="STYLE1">登录</h3>
      </div></td>
    </tr>
    <tr bgcolor="#F2E9ED">
      <td width="200" height="30"><div align="right"><strong>账号：</strong></div></td>
      <td width="300"><label>
        <input name="username" type="text" id="username" />
      </label></td>
    </tr>
    <tr bgcolor="#F2E9ED">
      <td height="30"><div align="right"><strong>密码：</strong></div></td>
      <td><label>
        <input type="password" name="password" id="password" />
      </label></td>
    </tr>
    <tr bgcolor="#F2E9ED">
      <td height="30" colspan="2"><label>
        <div align="center">
          <input name="button" type="submit" id="button" onclick="MM_validateForm('username','','RisNum','password','','RisNum');return document.MM_returnValue" value="登录" />
          <input type="reset" name="button2" id="button2" value="重置" />
        </div>
      </label>
        <label>
        <div align="center"></div>
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
