<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Index
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Index
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		IF Cbool(KSUser.UserLoginChecked)=True Then
		 Response.Redirect("../")
		End If
		%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head>
<title>会员登录-<%=KS.Setting(1)%></title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css" />
<style>
	/*全局样式*/
* { margin:0; padding:0; }
body { font:12px/20px Verdana; color:#666; text-align:center; background:#fff; }
ul { list-style:none; }
img { border:none; }
img, input, select, button { vertical-align:middle; color:#666 }
input, select { font:12px Verdana; }
button { cursor:pointer; }
optgroup option { padding-left: 15px;}
a{color:#597D7D}

/*表单元素*/
	/*输入框*/
.ipt_tx, .ipt_tx2 { border:1px solid #D2D2D2; background:#fff; line-height:16px; height:16px; padding:2px; margin:0; margin-left:2px}
.ipt_tx2 { border-color:#9DB5CA; background:#F2F9FE; }
	/*按钮*/
.btn { background:no-repeat; width:93px; height:28px; color:#333; font-size:12px; line-height:28px; border:none; }
.bnormal { background-image:url(../images/btn_normal.gif); }

.mainbody{width:100%; margin:0px auto;}
.header{width:980px; margin:0px auto;height:60px; margin-top:20px;}
.logo{width:215px;float:left;}
.txz{font-size:28px;font-weight:bold;color:#ff6600;float:left;line-height:56px;}
.other{width:300px;float:right;}
.other li{float:left;width:60px;text-align:center;line-height:56px;}
.controlBox{width:980px;margin:0 auto;background:url(../images/banner.png) no-repeat right 12px;}


.left{width:600px; float:left;}
.left .welcome{margin-top:130px;text-align:left;}
.left .welcome dt{font-size:14px;font-weight:bold;color:#ff6600;line-height:25px;}
.left .welcome dd{color:#000;line-height:25px;}



.right{width:314px; height:410px; float:right; margin-right:60px; display:inline; background:url(../images/login_box.png) no-repeat;}
.right strong{height:40px; display:block; text-indent:-10000px; overflow:hidden;}
/*after login*/
.right .box1{margin:20px 0 0 30px; font-size:14px;}
.right .box1 p{font:bold 14px "宋体";}
/*before login*/
.right .box1 form table{font-weight:bold;}
.right .box1 form table tr td{line-height:35px;}
.box1 .form_detail{text-align:left;font-size:14px;font-weight:bold;color:#000;line-height:28px;}
.box1 .form_detail p{margin:6px;}
.box1 .ipt_tx{padding:5px;line-height:20px;}
.box1 h2{font-size:14px;font-weight:700;color:#000; text-align:left;margin-top:30px;}
#showVerify{font-size:12px;color:red;}


#ft{color:#999; border-top:#f1f1f1 1px solid;}
#ft p{margin-top:9px;}
</style>
<script src="../../ks_inc/jquery.js"></script>
<script type="text/javascript">
var check={
   getCode:function(){
    $("#showVerify").html("<img align='absmiddle' src='../../plus/verifycode.asp' onClick='this.src=\"../../plus/verifycode.asp?n=\"+ Math.random();'>");
   },
   CheckForm:function(){
	 var username=$('#Username').val();
	 var pass=$('#Password').val();
	 var vycode=$('#Verifycode').val();
	 if (username==''){
		 alert('请输入用户名!');
		  $('#Username').focus();
		  return false;
	 }
	 if (pass==''){
		  alert('请输入登录密码!');
		  $('#Password').focus();
	      return false;
	 }
	 <%if KS.Setting(34)="1" then%>
	 if (vycode==''){
		  alert('请输入验证码!');
		  $('#Verifycode').focus();
		  return false;
	 }
	 <%end if%>
	}
}
</script>
</head>

<body>
<!-- head begin -->
<div class="mainbody">

	<div class="header">
		<div class="logo"><img alt="KesionCMS-通行证" src="../../images/logo.jpg"> </div>
		<div class="txz">|通行证</div>
		<div class="other">
		   <ul>
		     <li><a href="../../">首页</a></li> <li><a href="../user_Contributor.asp">匿名投稿</a></li><li><a href="../login" target="main">登录</a></li><li><a href="http://bbs.kesion.com" target="main">帮助</a></li>
		   </ul>
		</div>
	</div>
</div>
<!-- head end -->

<div class="controlBox">
	<div class="left">
            <div class="welcome">
			   <dl>
				<dt>个人信息</dt>
				<dd>登录会员中心后，您可以完善你的个人资料，设置安全问题，实时了解个人账户情况。</dd>
			   </dl>
			   <br />
			   <dl>
				<dt>个人/企业空间</dt>
				<dd>加入我们您将免费拥有一个空间，个人会员将得到一个个人空间,您可以在上面写日志、上传照片、找朋友、加入圈子讨论等。企业会员将得到一个企业空间，您可以将公司的简介、公司产品、公司动态、公告招聘信息等发布到您的空间。</dd>
			   </dl>
			   			   <br />

			   <dl>
				<dt>求职招聘</dt>
				<dd>您可以在会员中心发布求职信息，招聘信息；个人简历，公司介绍等。</dd>
			   </dl>
			</div>
            
        </div>
        <div class="right">
        
      <strong>登录通行证</strong>
            <div class="box1">
                <form action="../CheckUserLogin.asp" id="myform" name="myform" method="post">
				<div class="form_detail">
					<p>
						<label>用户名：</label>
						<input type="text" name="Username" maxlength="60" id="Username" class="ipt_tx" style="width:149px;" tabindex="1" />
						
					</p>
					<p>
						<label>密&nbsp;&nbsp;码：</label>
						<input type="password" name="Password" maxlength="60" id="Password" class="ipt_tx" style="width:149px;" tabindex="2" autocomplete="off"/>
					</p>
					<%if KS.Setting(34)="1" then%>
					<p>
						<label>验证码：</label>
						<input type="text" maxlength="6" name="Verifycode" id="Verifycode" onFocus="this.value='';check.getCode()" class="ipt_tx" style="width:55px;" tabindex="3" autocomplete="off"/>
						<span id="showVerify">鼠标点击输入框显示</span>
						
					</p>
					<%end if%>
					<p><input type="hidden" name="u1" id="u1"/>
						<input type="submit" tabindex="5"  onClick="return check.CheckForm();" class="btn bnormal" value="登  录">
					<a tabindex="6" href="../GetPassword.asp" target="_blank">忘记密码？</a>  <a tabindex="7" href="../ActiveCode.asp" target='_blank'>重发激活码</a></p>
				</div>
			</form>
			<h2>还没有会员通行证帐号？</h2>
			<div class="form_detail">
				<p>
					<input type="button" tabindex="8" id="btn_regist" class="btn bnormal"  onclick="location.href='../../?do=reg'" value="现在就注册" />
				</p>
			</div>
			<h2>免费注册本站会员，仅需花费您30秒钟！</h2>

            </div>
           
		
        </div>
        <div class="clear"></div>
    </div>

  <br />

	<div id="ft">
		<p>漳州市科兴信息技术有限公司 &copy; 版权所有</p>
		<p>网址：http://www.kesion.com QQ：9537636 41904294</p>
	</div>
</body>
</html>
        <%
  End Sub
End Class
%> 
