<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312" 
Dim KSCls
Set KSCls = New UserLogin
KSCls.Kesion()
Set KSCls = Nothing

Class UserLogin
        Private KS
		Private KSUser,Action
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		 Action=KS.S("Action")
		 If Action="script" Then
		  Call GetLoginByScript()
		  Exit Sub
		 ElseIf Action="checklogin" Then
		  Call CheckUserIsLogin()
		  Exit Sub
		 ElseIf Action="PoploginStr" Then
		  GetPoploginStr()
		  Exit Sub
		 End If
		%>
		<html>
<head>
<title>��Ա��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.textbox
{
BACKGROUND-COLOR: #ffffff;
BORDER: #ccc 1px solid;
COLOR: #999;
HEIGHT: 22px;
line-height:22px
border-color: #666666 #666666 #666666 #666666; font-size: 9pt;FONT-FAMILY: verdana
}
TD
{
FONT-FAMILY:����;FONT-SIZE: 9pt;line-height: 130%;
}
a{text-decoration: none;} /* �������»���,��Ϊunderline */
a:link {color: #000000;} /* δ���ʵ����� */
a:visited {color: #333333;} /* �ѷ��ʵ����� */
a:hover{COLOR: #AE0927;} /* ����������� */
a:active {color: #0000ff;} /* ����������� */
.logintitle{font-size:14px;color:#336699;font-weight:bold}
#PopLogin td{font-size:14px;line-height:180%}
#PopLogin td a{color:#336699;text-decoration:underline}
#PopLogin td span{color:#5F5C67;font-size:13px}
#PopLogin td input{margin:2px}
.btn{border-color:#3366cc;margin-right:1em;color:#fff;background:#3366cc;}
.btn{border-width:1px;cursor:pointer;padding:.1em 1em;*padding:0 1em;font-size:9pt; line-height:130%; overflow:visible;}
-->
</style>
<script language="javascript">
//if(self==top){self.location.href="index.asp";}
function CheckForm(){
	var username=document.myform.Username.value;
	var pass=document.myform.Password.value;
	if (username=='')
	{
	  alert('�������û���!');
	  document.myform.Username.focus();
	  return false;
    }
	if (pass=='')
	{
	  alert('�������¼����!');
	  document.myform.Password.focus();
	  return false;
	 }
	 <% If KS.Setting(34)="1" Then%>
	 if (document.myform.Verifycode.value==''){
	  alert('��������֤��!');
	  return false;
	 }
	 <%End If%>
	 return true;
}
function getCode(){
 document.getElementById('showVerify').innerHTML='<IMG style="cursor:pointer" src="../plus/verifycode.asp?n="<%=Timer%>" onClick="this.src=\'../plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">';
}
</script>

</head>
<body leftmargin="0" topmargin="0" style="background-color:transparent;<%If KS.S("Action")="Poplogin" then response.write "background:url(images/loginbg.png) repeat-x;"%>">
		<%
		If KS.S("Action")="Top" Then
		   Call Login1()
		ElseIf KS.S("Action")="Poplogin" Then
		   Call PopLogin()
		Else
		   Call Login2()
		 End If
		End Sub
		
		'script��ʽ����
		Sub GetLoginByScript()
           If KSUser.UserLoginChecked=false Then
		    KS.Echo "document.write('<form name=""myform"" id=""myform"" method=""POST"" action=""" & KS.GetDomain & "user/checkuserlogin.asp"">�û��� <input type=""text"" maxlength=""30"" name=""username"" id=""username"" size=""12"" class=""textbox""/>&nbsp;���� <input style=""FONT-FAMILY: verdana;"" type=""password"" maxlength=""30"" name=""password"" size=""12"" id=""password"" class=""textbox""/>&nbsp;');"
			 If KS.Setting(34)="1" Then
			  KS.Echo "document.write('<span>��֤�� </span><input onFocus=""getCode()"" maxlength=""8"" type=""text"" name=""Verifycode"" size=""5"" class=""textbox""><span id=""showVerify""><IMG style=""cursor:pointer"" src=""" & KS.GetDomain & "plus/verifycode.asp"" onClick=""this.src=\'" & KS.GetDomain & "plus/verifycode.asp?n=\'+ Math.random();"" align=""absmiddle""></span>');"
			 End If
			 KS.Echo "document.write('<input align=""absmiddle"" type=""image"" src=""" & KS.GetDomain & "images/login.gif"" onclick=""return(CheckLoginForm())""  class=""lgbtn""/>&nbsp;<a href=""" & KS.GetDomain & "?do=reg"" target=""_self"">ע��</a>&nbsp;|&nbsp;<a href=""" & KS.GetDomain & "user/getpassword.asp"">�һ�����</a></form>');"
		   Else
		     KS.Echo "document.write('���ã�<span style=""color:red"">" & KSUser.UserName & "</span>,��ӭ������Ա����!��<a href=""" & KS.GetDomain & "user/"">��Ա����</a>����<a href=""" & KS.GetDomain & "/user/user_Message.asp?action=inbox"">����Ϣ"& GetMailTips()& "</a>����<a href=""" & KS.GetDomain & "User/UserLogout.asp"">�˳�</a>��');"
		   End If
		End Sub
		
		Sub CheckUserIsLogin()
		  If KSUser.UserLoginChecked=false Then
		    If KS.S("S")="1" Then
		     KS.Echo "var user={'loginstr':'����,��ӭ����" & KS.Setting(0) & "! [<a href=""javascript:void(0)"" onclick=""ShowPopLogin()"">��¼</a>] | [<a href=""" & KS.GetDomain & "?do=reg"" target=""_blank"">���ע��</a>]'}"
			Else
			 KS.Echo "var user={'loginstr':'<form name=""myform"" id=""myform"" method=""POST"" action=""" & KS.GetDomain & "user/checkuserlogin.asp"">�û��� <input type=""text"" maxlength=""30"" name=""username"" id=""username"" size=""12"" class=""textbox""/>&nbsp;���� <input style=""FONT-FAMILY: verdana;"" type=""password"" maxlength=""30"" name=""password"" size=""12"" id=""password"" class=""textbox""/>&nbsp;"
			 If KS.Setting(34)="1" Then
			  KS.Echo "<span>��֤�� </span><input onFocus=""getCode()"" maxlength=""8"" type=""text"" name=""Verifycode"" size=""5"" class=""textbox""><span id=""showVerify""><IMG style=""cursor:pointer"" src=""" & KS.GetDomain & "plus/verifycode.asp"" onClick=""this.src=\'" & KS.GetDomain & "plus/verifycode.asp?n=\'+ Math.random();"" align=""absmiddle""></span>"
			 End If
			 KS.Echo "<input align=""absmiddle"" type=""image"" src=""" & KS.GetDomain & "images/login.gif"" onclick=""return(CheckLoginForm())""  class=""lgbtn""/>&nbsp;<a href=""" & KS.GetDomain & "?do=reg"" target=""_self"">ע��</a>&nbsp;|&nbsp;<a href=""" & KS.GetDomain & "user/getpassword.asp"">�һ�����</a></form>'}"
			End If
		  Else
		    KS.Echo "var user={'loginstr':'���ã�<span style=""color:red"">" & KSUser.UserName & "</span>,��ӭ������Ա����!��<a href=""" & KS.GetDomain & "user/"">��Ա����</a>����<a href=""" & KS.GetDomain & "/user/user_Message.asp?action=inbox"">����Ϣ"& GetMailTips()& "</a>����<a href=""" & KS.GetDomain & "User/UserLogout.asp"">�˳�</a>��'}"
		  End If
		End Sub
		
		
		Function GetMailTips()
		    Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
			'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogMessage Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogComment Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_Friend Where Friend='" &KSUser.UserName &"' And accepted=0")(0)
			 IF MyMailTotal>0 Then 
			  GetMailTips="<font color=""red"">" & MyMailTotal & "</font><bgsound src=""" & KS.GetDomain & "User/images/mail.wav"" border=0>"  
			  Else
			  GetMailTips=0
			 End If
		End Function
		
		Sub PopLogin()
		%>
		 <table id="PopLogin" width="100%" height="184" cellpadding="0" cellspacing="0" border="0">
		  <tr>
		   <td>
		     <table border="0" width="95%" align="center">
			   <form action="checkuserlogin.asp" method="post" name="myform">
			  <tr>
			   <td style="border-right:solid 1px #cccccc">
			    û���˺ţ�<a href="../?do=reg" target="_blank">����ע��</a><br/>
				��������, <a href="<%=KS.Setting(3)%>user/getpassword.asp"  target="_blank">��Ҫ�һ�</a> <br />
			   </td>
			   <td>
			      <div class="logintitle">�û���¼</div>
				  <span>�û��˺ţ�</span><input type="text" name="Username" class="textbox"><br />
				  <span>��¼���룺</span><input type="password" name="Password" class="textbox"><br/>
				  <% If KS.Setting(34)="1" Then%>
				  <span>�����ַ���</span><input onFocus="getCode()" type="text" name="Verifycode" size="5" class="textbox"><span id='showVerify'></span><br/>
				  <%end if%>
				  <input type="hidden" name="Action" value="PopLogin">
				  <input type="submit" value=" �� ¼ " class="btn" onClick="return(CheckForm())"  name="submit">
				   <input name="ExpiresDate" type="checkbox" id="ExpiresDate" value="checkbox">	<span>���õ�¼</span>
			   </td>
			  </tr>
				 </form>
			 </table>
		   </td>
		  </tr>
		 </table>
		<%
		End Sub
		
		'��$.getScript����������
		Sub GetPoploginStr()
		 Dim Str
		 str="<table id=""PopLogin"" style=""font-size:14px;line-height:180%"" width=""100%"" height=""184"" cellpadding=""0"" cellspacing=""0"" border=""0""><tr><td><table border=""0"" width=""95%"" align=""center""><tr><td style=""border-right:solid 1px #cccccc"">û���˺ţ�<a href=""" & KS.GetDomain & "?do=reg"" target=""_blank"">����ע��</a><br/>��������, <a href=""" & KS.GetDomain &"user/getpassword.asp""  target=""_blank"">��Ҫ�һ�</a> <br /></td><td style=""text-align:left""><div style=""font-size:14px;color:#336699;font-weight:bold"" class=""logintitle"">�û���¼</div><span>�û��˺ţ�</span><input type=""text"" name=""Username"" class=""textbox""><br /><span>��¼���룺</span><input type=""password"" name=""Password"" class=""textbox""><br/>"
		 If KS.Setting(34)="1" Then
		 str=str &"<span>�����ַ���</span><input onFocus=""getCode()"" type=""text"" name=""Verifycode"" size=""5"" class=""textbox""><span id=""showVerify""></span><br/>"
		 End If
		 Str=Str & "<input type=""submit"" onclick=""return(CheckLoginForm())"" value="" �� ¼ "" name=""submit""><input name=""ExpiresDate"" type=""checkbox"" id=""ExpiresDate"" value=""checkbox"">	<span>���õ�¼</span></td></tr></table></td></tr></table>"
         KS.Die "var userpop={""str"":'" & str & "'}"
		End Sub
		
		Sub Login1()
			If KSUser.UserLoginChecked() = False Then
			%>
				<table cellspacing="0" cellpadding="0" width="99%" border="0">
				<form name="myform" action="<%=KS.GetDomain%>User/CheckUserLogin.asp?Action=Top" method="post">
								<tr>
								  <td>�û�����<input class="textbox" size="10" name="Username" />  �� �룺<input class="textbox" type="Password" size="10" name="Password"><%if KS.Setting(34)=1 Then%>��֤�룺<input name="Verifycode" type="text" class="textbox" id="Verifycode" size="6" /><%
				Response.Write "<IMG style=""cursor:pointer"" src=""" & KS.GetDomain & "plus/verifycode.asp?n=" & Timer & """ onClick=""this.src='" & KS.GetDomain & "plus/verifycode.asp?n='+ Math.random();"" align=""absmiddle"">"
				end if%> 
								    <input name="loginsubmit" type="image"  onClick="return(CheckForm())" src="<%=KS.GetDomain%>images/login.gif" align="top" />
								    &nbsp;<a href="<%=KS.GetDomain%>?do=reg" target="_blank"><img src="<%=KS.GetDomain%>images/reg.gif"  border="0" align="absmiddle" twffan="done" /></a></td>
								</tr>
							</table>
			<%Else

			%>
			<table cellspacing="0" cellpadding="0" width="99%" border="0">
				<tr>
			     <td height="22" align="center">����!<font color=red><%=KSUser.UserName%></font>,��ӭ������Ա����!&nbsp;��<a href="<%=KS.GetDomain%>User/index.asp?User_Message.asp?action=inbox" target="_parent">������ <%=GetMailTips()%></a>��&nbsp;��<a href="<%=KS.GetDomain%>User/index.asp" target="_parent">��Ա����</a>��&nbsp;��<a href="<%=KS.GetDomain%>User/UserLogout.asp">�˳���¼</a>��</td>
				</tr>
			</table>
<%End IF
		End Sub
		Sub Login2()
			If KSUser.UserLoginChecked() = False Then
			%>
			<table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
			 <form name="myform" action="CheckUserLogin.asp" method="post">
			  <tr>
				<td height="25">�û�����
				<input name="Username" type="text" class="textbox" id="Username" size="15"></td>
			  </tr>
			  <tr>
				<td height="25">�ܡ��룺
				<input name="Password" type="password" class="textbox" id="Password" size="16"></td>
			  </tr>
			  <%if KS.Setting(34)=1 Then%>
			  <tr>
				<td height="25">��֤�룺
				<input name="Verifycode" onClick="getCode()" type="text" class="textbox" id="Verifycode" size="6">
				<span id='showVerify'></span>
				</td>
			  </tr>
			  <%end if%>
			  <tr>
				<td height="25"><div align="center"><img src="<%=KS.GetDomain%>images/losspass.gif" align="absmiddle"> <a href="<%=KS.GetDomain%>User/GetPassword.asp" target="_parent">��������</a> <img src="<%=KS.GetDomain%>images/mas.gif" align="absmiddle"> <a href="<%=KS.GetDomain%>?do=reg" target="_parent">�»�Աע��</a>    </div></td>
			  </tr>
			  <tr>
				<td height="25"><div align="center">
				  <input type="submit" name="Submit"  onClick="return(CheckForm())" class="inputButton" value="��¼">

				  <input name="ExpiresDate" type="checkbox" id="ExpiresDate" value="checkbox">
			���õ�¼</div></td>
			  </tr>
			  </form>
            </table>
			<%Else
			 dim  ChargeTypeStr
			 if KSUser.ChargeType=1 Then
			   ChargeTypeStr="�۵�"
			 elseif KSUser.ChargeType=2 Then
			   ChargeTypeStr="��Ч��"
			 else
			   ChargeTypeStr="������"
			 End If
			%>
			<table align="center" style="margin-top:5px" width="80%" border="0" cellspacing="0" cellpadding="0">
			<tr><td align="center"><font color=red><%=KSUser.UserName%></font>,
           <%
			If (Hour(Now) < 6) Then
            Response.Write "<font color=##0066FF>�賿��!</font>"
			ElseIf (Hour(Now) < 9) Then
				Response.Write "<font color=##000099>���Ϻ�!</font>"
			ElseIf (Hour(Now) < 12) Then
				Response.Write "<font color=##FF6699>�����!</font>"
			ElseIf (Hour(Now) < 14) Then
				Response.Write "<font color=##FF6600>�����!</font>"
			ElseIf (Hour(Now) < 17) Then
				Response.Write "<font color=##FF00FF>�����!</font>"
			ElseIf (Hour(Now) < 18) Then
				Response.Write "<font color=##0033FF>�����!</font>"
			Else
				Response.Write "<font color=##ff0000>���Ϻ�!</font>"
			End If
			%>&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>�Ʒѷ�ʽ�� <strong><%= ChargeTypeStr%></strong> </td></tr>
			<tr><td>������֣� <strong><%=KSUser.GetUserInfo("Score")%></strong> ��</td></tr>
			<%if KSUser.ChargeType=1 or KSUser.ChargeType=2 then%>
			<% if KSUser.ChargeType=1 then%>
			<tr><td>���õ�ȯ�� <strong><%=KSUser.GetUserInfo("Point")%></strong> ��</td></tr>
			<%else%>
			<tr><td>ʣ�������� <strong><%=KSUser.GetEdays%></strong></td></tr>
			<%end if%>
			<%end if%>
			<tr><td>���Ķ��ţ� <strong><%=GetMailTips()%></strong> ��</td></tr>
			<tr><td>��¼������ <strong><%=KSUser.GetUserInfo("LoginTimes")%></strong> ��</td></tr>
            <tr><td nowrap="nowrap">��<a href="<%=KS.GetDomain%>User/index.asp" target="_parent">��Ա����</a>����<a href="<%=KS.GetDomain%>User/UserLogout.asp">�˳���¼</a>��</td></tr>
			</table>
<%End IF
  End Sub
End Class
%>

 
