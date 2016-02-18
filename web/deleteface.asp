<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/teach.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_teach_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Score" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
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
    <h3>删除学生成绩</h3>
    <p>在文本框内输入需要修改的学生学号</p>
  </div>
</div>
<form id="form1" name="form1" method="post" action="/deleteScore.asp">
  <table width="400" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#F2E9ED">
    <tr>
      <td align="center"><label>
      <select name="studentID" id="studentID">
        <%
While (NOT Recordset1.EOF)
%>
        <option value="<%=(Recordset1.Fields.Item("studentID").Value)%>"><%=(Recordset1.Fields.Item("studentID").Value)%></option>
        <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
      </select>
      </label></td>
      <td><label>
        <input name="button" type="submit" id="button" onclick="MM_validateForm('studentID','','RisNum');return document.MM_returnValue" value="删除" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
