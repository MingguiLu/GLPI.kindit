<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim KS:Set KS=New PublicCls

Select Case  KS.G("Action")
 Case "LoginCheck"
  Call CheckLogin()
 Case "LoginOut"
  Call LoginOut()
 Case Else
  Call CheckSetting()
  Call Main()
End Select

Sub CheckSetting()
     dim strDir,strAdminDir,InstallDir
	 strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
	 strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
	 InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			
	If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
	   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
	End If
 If KS.Setting(2)<>KS.GetAutoDoMain or KS.Setting(3)<>InstallDir Then
	
  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open "Select Setting From KS_Config",conn,1,3
  Dim SetArr,SetStr,I
  SetArr=Split(RS(0),"^%^")
  For I=0 To Ubound(SetArr)
   If I=0 Then 
    SetStr=SetArr(0)
   ElseIf I=2 Then
    SetStr=SetStr & "^%^" & KS.GetAutoDomain
   ElseIf I=3 Then
    SetStr=SetStr & "^%^" & InstallDir
   Else
    SetStr=SetStr & "^%^" & SetArr(I)
   End If
  Next
  RS(0)=SetStr
  RS.Update
  RS.Close:Set RS=Nothing
  Call KS.DelCahe(KS.SiteSn & "_Config")
  Call KS.DelCahe(KS.SiteSn & "_Date")
 End If
End Sub

Sub Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE><%=KS.Setting(0) & "---网站后台管理"%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta http-equiv="X-UA-Compatible" content="IE=7" />
<script language="JavaScript" type="text/JavaScript" src="Include/SoftKeyBoard.js"></script>

<STYLE>
body{ margin:0px auto; padding:0px auto; font-size:12px; color:#555; font-family:Verdana, Arial, Helvetica, sans-serif;text-align:center; background:#016AA9 url(images/body_bg.jpg) no-repeat center top;}
html {
overflow:hidden;
} 
form,div,ul,li {margin:0px auto;padding: 0; border: 0; }
img,a img{border:0; margin:0; padding:0;}
a{font-size:12px;color:#000000}
a:link,a:visited{color:#00436D;}
td {
	font-size: 12px; color: #fff; LINE-HEIGHT: 20px; FONT-FAMILY: "", Arial, Tahoma
}
.textbox{height:16px;}
.head{height:74px; background:url(images/headbg.jpg);}
.header{width:980px; margin:0px auto;}
.logo{width:300px;height:74px; background:url(images/logo.jpg) no-repeat;float:left;margin:0px;padding:0px;}
.right{width:580px;float:right; text-align:right;color:#fff;font-family:"宋体"; line-height:68px; }
.right a:link,.right a:visited{color:#fff; text-decoration: none}
.right a:hover{color:#FFFC06; text-decoration:underline}
.main{margin-top:68px;width:100%;text-align:center}
.login{width:667px;height:315px;background: url('images/loginbg.jpg') no-repeat left top;
}
.form{
 width:360px;
 float:right;
 padding-top:120px;
 bottom:0px;
 text-align:left;
}
.form .user{
height:34px;
background:url('images/login.gif') no-repeat 0px 0px;
width:230px;
padding-left:18px;
margin-left:13px;

}
.form .password{
height:34px;
background:url('images/login.gif') no-repeat 0px -40px;
width:230px;
padding-left:18px;
margin-left:13px;
}
.form .Verifycode{
height:34px;
background:url('images/login.gif') no-repeat 0px -80px;
width:235px;
padding-left:18px;
margin-left:13px;
}

.form .AdminLoginCode{
height:34px;
background:url('images/login.gif') no-repeat 0px -120px;
width:235px;
padding-left:18px;
margin-left:13px;
}

.form .user input{margin-left:59px;*+margin-left:48px;_marign-left:48px;margin-top:6px;border:0px;background:none;}
#login{text-align:left;padding-left:6px;}
.form .password input{FONT-FAMILY: verdana;margin-left:65px;*+margin-left:53px;_margin-left:53px;margin-top:4px;border:0px;background:none;}
.form .Verifycode input{width:74px;margin-left:33px;*+margin-left:16px;_margin-left:15px;margin-top:6px;border:0px;background:none;}
.form .Verifycode img{margin-left:10px;margin-top:-5px;*+margin-top:5px;*+margin-left:1px;_margin-left:2px;}
.form .AdminLoginCode input{FONT-FAMILY: verdana;width:74px;margin-left:-21px;*+margin-left:-30px;_margin-left:-30px;margin-top:6px;border:0px;background:none;}

.form .left{width:255px;*+width:255px;_width:235px;float:left;text-align:center}
.form .right{width:88px;float:right;text-align:center;}
.form .btn{cursor:pointer;margin-right:5px;border:0px;width:76px;height:66px;background:url(images/login.gif) no-repeat -252px 0px;}
.form .btn:hover{background:url(images/login.gif) no-repeat -252px -66px;}

#copyright {margin:0px auto;left:250px;margin-top:100px;bottom:100px;text-align:center;width:550px;PADDING: 1px 1px 1px 1px; FONT: 12px 
verdana,arial,helvetica,sans-serif; COLOR: #fff; TEXT-DECORATION: none}
#copyright a:link,#copyright a:visited{color:#FFEA04; font-size:12px;}
</STYLE>
</head>
<body>
<div class="head">
 <div class="header">
  <div class="logo"></div>
  <div class="right"><a href="http://www.kesion.com" target="_blank">官方首页</a> | <a href="http://help.kesion.com" target="_blank"> 帮助中心</a> | <a href="http://www.kesion.com" target="_blank"> 会员中心</a> | <a href="http://bbs.kesion.com" target="_blank"> 交流论坛</a></div>
 </div>
</div>
<div class="main">
 <FORM ACTION="Login.asp?Action=LoginCheck" method="post" name="LoginForm" onSubmit="return(CheckForm(this))">
    <div class="login">
	    <div class="form">
		
		<div class="left">
		 <div class="user">
		 <input type="text" maxlength="50" tabindex="1" name="UserName" id="UserName"/>
		 </div>
		 
		 <div class="password">
		 <%IF KS.Setting(98)=1 Then%>
		 <input name="PWD" type="password"  tabindex="2" onFocus="this.select();" onChange="Calc.password.value=this.value;" onClick="password1=this;showkeyboard();this.readOnly=1;Calc.password.value=''" onKeyDown="Calc.password.value=this.value;" maxlength="50" readOnly>
         <%Else%>
		 <input type="password" tabindex="2" maxlength="50" name="PWD" id="textbox"/>
		 <%End If%>
				
		 </div>

		 <div class="Verifycode">
		 <input type="text" maxlength="6" tabindex="3" name="Verifycode" id="textbox"/>
		<IMG style="cursor:pointer;" src="../plus/verifycode.asp?n=<%=Timer%>" onClick="this.src='../plus/verifycode.asp?n='+ Math.random();" align="absmiddle"> </div>

        <%if EnableSiteManageCode = True Then%>
		 <div class="AdminLoginCode">
		 <input type="password" maxlength="20" tabindex="4" name="AdminLoginCode" id="textbox"/>
		 </div>
		  <%
		End If
	  if EnableSiteManageCode=true And SiteManageCode="8888" Then
	   Response.Write"<br /><span style='color:#086898'>原始认证码为<span style='color:#ff0000'>8888</span>,可打开conn.asp修改</span>"
	  End If
	  %>
		</div>
		
		<div class="right">

            <input type="submit" tabindex="5" class="btn" value=" ">
	    </div>

		</div>
    </div>
 </form>
</div>
<br />
<div class="line"></div>
<div class="botinfo" id="copyright"> 
漳州科兴信息技术有限公司 Copyright &copy;2006-2011 <a href="http://www.kesion.com" target="_blank"> www.kesion.com</a>,All Rights Reserved. </div>


<script type="text/javascript">
<!--
function document.onreadystatechange()
{  var app=navigator.appName;
  var verstr=navigator.appVersion;
  if(app.indexOf('Netscape') != -1) {
    alert('友情提示：\n    您使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。');
  } else if(app.indexOf('Microsoft') != -1) {
    if (verstr.indexOf('MSIE 3.0')!=-1 || verstr.indexOf('MSIE 4.0') != -1 || verstr.indexOf('MSIE 5.0') != -1 || verstr.indexOf('MSIE 5.1') != -1)
      alert('友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。');
  }
  document.LoginForm.UserName.focus();
}
function CheckForm(ObjForm) {
  if(ObjForm.UserName.value == '') {
    alert('请输入管理账号！');
    ObjForm.UserName.focus();
    return false;
  }
  if(ObjForm.PWD.value == '') {
    alert('请输入授权密码！');
    ObjForm.PWD.focus();
    return false;
  }
  if (ObjForm.PWD.value.length<6)
  {
    alert('授权密码不能少于六位！');
    ObjForm.PWD.focus();
    return false;
  }
  if (ObjForm.Verifycode.value == '') {
    alert ('请输入验证码！');
    ObjForm.Verifycode.focus();
    return false;
  }
  <%if EnableSiteManageCode = True Then%>
  if (ObjForm.AdminLoginCode.value == '') {
    alert ('请输入后台管理认证码！');
    ObjForm.AdminLoginCode.focus();
    return false;
  }
  <%End If%>
}
//-->
</script>
</html>
<%End Sub
Sub CheckLogin()
  Dim PWD,UserName,LoginRS,SqlStr,RndPassword
  Dim ScriptName,AdminLoginCode
  AdminLoginCode=KS.G("AdminLoginCode")
  IF Trim(Request.Form("Verifycode"))<>Trim(Session("Verifycode")) then 
   Call KS.Alert("登录失败:\n\n验证码有误，请重新输入！","Login.asp")
   exit Sub
  end if
  If EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode Then
   Call KS.Alert("登录失败:\n\n您输入的后台管理认证码不对，请重新输入！","Login.asp")
   exit Sub
  End If
  Pwd =MD5(KS.R(Request.form("pwd")),16)

  UserName = KS.R(trim(Request.form("username")))
  RndPassword=KS.R(KS.MakeRandomChar(20))
  ScriptName=KS.R(Trim(Request.ServerVariables("HTTP_REFERER")))
  
  Set LoginRS = Server.CreateObject("ADODB.RecordSet")
  SqlStr = "select * from KS_Admin where UserName='" & UserName & "'"
  LoginRS.Open SqlStr,Conn,1,3
  If LoginRS.EOF AND LoginRS.BOF Then
	  Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的帐号!")
      Call KS.AlertHistory("登录失败:\n\n您输入了错误的帐号，请再次输入！",-1)
  Else
  
     IF LoginRS("PassWord")=pwd THEN
       IF Cint(LoginRS("Locked"))=1 Then
          Call KS.Alert("登录失败:\n\n您的账号已被管理员锁定，请与您的系统管理员联系！","Login.asp")	
	      Response.End
	   Else
		  	 '登录成功，进行前台验证，并更新数据
			  Dim UserRS:Set UserRS=Server.CreateObject("Adodb.Recordset")
			  UserRS.Open "Select top 1 Score,LastLoginIP,LastLoginTime,LoginTimes,UserName,Password,RndPassWord,IsOnline,UserID From KS_User Where UserName='" & LoginRS("PrUserName") & "' and GroupID=1",Conn,1,3
			  IF Not UserRS.Eof Then
			  
						If datediff("n",UserRS("LastLoginTime"),now)>=KS.Setting(36) then '判断时间
						UserRS("Score")=UserRS("Score")+KS.Setting(37)
						end if
					 UserRS("LastLoginIP") = KS.GetIP
					 UserRS("LastLoginTime") = Now()
					 UserRS("LoginTimes") = UserRS("LoginTimes") + 1
					 UserRS("RndPassWord") = RndPassWord
					 UserRS("IsOnline")=1
					 UserRS.Update		
	
					'置前台会员登录状态
                    If EnabledSubDomain Then
							Response.Cookies(KS.SiteSn).domain=RootDomain					
					Else
                            Response.Cookies(KS.SiteSn).path = "/"
					End If		
					 Response.Cookies(KS.SiteSn)("UserID") = UserRS("UserID")
					 Response.Cookies(KS.SiteSn)("UserName") = KS.R(UserRS("UserName"))
			         Response.Cookies(KS.SiteSn)("Password") = UserRS("Password")
					 Response.Cookies(KS.SiteSn)("RndPassword") = KS.R(UserRS("RndPassword"))
					 Response.Cookies(KS.SiteSn)("AdminLoginCode") = AdminLoginCode
					 Response.Cookies(KS.SiteSn)("AdminName") =  UserName
					 Response.Cookies(KS.SiteSn)("AdminPass") = pwd
					 Response.Cookies(KS.SiteSn)("SuperTF")   = LoginRS("SuperTF")
					 Response.Cookies(KS.SiteSn)("PowerList") = LoginRS("PowerList")
					 Response.Cookies(KS.SiteSn)("ModelPower") = LoginRS("ModelPower")
					 'Response.Cookies(KS.SiteSn).Expires = DateAdd("h", 3, Now())   '3小时没有操作自动失败
             Else 
				   Call KS.InsertLog(UserName,0,ScriptName,"找不到前台账号!")
				   Call KS.Alert("登录失败:\n\n找不到前台账号！","Login.asp")	
				   Response.End
			 End If
			   UserRS.Close:Set UserRS=Nothing
			   
	  LoginRS("LastLoginTime")=Now
	  LoginRS("LastLoginIP")=KS.GetIP
	  LoginRS("LoginTimes")=LoginRS("LoginTimes")+1
	  LoginRS.UpDate
	  Call KS.InsertLog(UserName,1,ScriptName,"成功登录后台系统!")
	   Response.Redirect("Index.asp")
	End IF
  ELse
     If EnabledSubDomain Then
		Response.Cookies(KS.SiteSn).domain=RootDomain					
	 Else
        Response.Cookies(KS.SiteSn).path = "/"
	End If
    Response.Cookies(KS.SiteSn)("AdminName")=""
	Response.Cookies(KS.SiteSn)("AdminPass")=""
	Response.Cookies(KS.SiteSn)("SuperTF")=""
	Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
	Response.Cookies(KS.SiteSn)("PowerList")=""
	Response.Cookies(KS.SiteSn)("ModelPower")=""
	Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的口令:" & Request.form("pwd"))
    Call KS.Alert("登录失败:\n\n您输入了错误的口令，请再次输入！","Login.asp")	
  END IF
 End If
END Sub
Sub LoginOut()
		   Conn.Execute("Update KS_Admin Set LastLogoutTime=" & SqlNowString & " where UserName='" & KS.R(KS.C("AdminName")) &"'")
		   Dim AdminDir:AdminDir=KS.Setting(89)
		   If EnabledSubDomain Then
				Response.Cookies(KS.SiteSn).domain=RootDomain					
			Else
                Response.Cookies(KS.SiteSn).path = "/"
			End If
			Response.Cookies(KS.SiteSn)("PowerList")=""
			Response.Cookies(KS.SiteSn)("AdminName")=""
			Response.Cookies(KS.SiteSn)("AdminPass")=""
			Response.Cookies(KS.SiteSn)("SuperTF")=""
			Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
			Response.Cookies(KS.SiteSn)("ModelPower")=""
			session.Abandon()
			Response.Write ("<script> top.location.href='" & KS.Setting(2) & KS.Setting(3) &"';</script>")
End Sub
Set KS=Nothing
%>