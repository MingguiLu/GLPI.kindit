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
<title>��Ա��¼-<%=KS.Setting(1)%></title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css" />
<style>
	/*ȫ����ʽ*/
* { margin:0; padding:0; }
body { font:12px/20px Verdana; color:#666; text-align:center; background:#fff; }
ul { list-style:none; }
img { border:none; }
img, input, select, button { vertical-align:middle; color:#666 }
input, select { font:12px Verdana; }
button { cursor:pointer; }
optgroup option { padding-left: 15px;}
a{color:#597D7D}

/*��Ԫ��*/
	/*�����*/
.ipt_tx, .ipt_tx2 { border:1px solid #D2D2D2; background:#fff; line-height:16px; height:16px; padding:2px; margin:0; margin-left:2px}
.ipt_tx2 { border-color:#9DB5CA; background:#F2F9FE; }
	/*��ť*/
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
.right .box1 p{font:bold 14px "����";}
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
		 alert('�������û���!');
		  $('#Username').focus();
		  return false;
	 }
	 if (pass==''){
		  alert('�������¼����!');
		  $('#Password').focus();
	      return false;
	 }
	 <%if KS.Setting(34)="1" then%>
	 if (vycode==''){
		  alert('��������֤��!');
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
		<div class="logo"><img alt="KesionCMS-ͨ��֤" src="../../images/logo.jpg"> </div>
		<div class="txz">|ͨ��֤</div>
		<div class="other">
		   <ul>
		     <li><a href="../../">��ҳ</a></li> <li><a href="../user_Contributor.asp">����Ͷ��</a></li><li><a href="../login" target="main">��¼</a></li><li><a href="http://bbs.kesion.com" target="main">����</a></li>
		   </ul>
		</div>
	</div>
</div>
<!-- head end -->

<div class="controlBox">
	<div class="left">
            <div class="welcome">
			   <dl>
				<dt>������Ϣ</dt>
				<dd>��¼��Ա���ĺ�������������ĸ������ϣ����ð�ȫ���⣬ʵʱ�˽�����˻������</dd>
			   </dl>
			   <br />
			   <dl>
				<dt>����/��ҵ�ռ�</dt>
				<dd>���������������ӵ��һ���ռ䣬���˻�Ա���õ�һ�����˿ռ�,������������д��־���ϴ���Ƭ�������ѡ�����Ȧ�����۵ȡ���ҵ��Ա���õ�һ����ҵ�ռ䣬�����Խ���˾�ļ�顢��˾��Ʒ����˾��̬��������Ƹ��Ϣ�ȷ��������Ŀռ䡣</dd>
			   </dl>
			   			   <br />

			   <dl>
				<dt>��ְ��Ƹ</dt>
				<dd>�������ڻ�Ա���ķ�����ְ��Ϣ����Ƹ��Ϣ�����˼�������˾���ܵȡ�</dd>
			   </dl>
			</div>
            
        </div>
        <div class="right">
        
      <strong>��¼ͨ��֤</strong>
            <div class="box1">
                <form action="../CheckUserLogin.asp" id="myform" name="myform" method="post">
				<div class="form_detail">
					<p>
						<label>�û�����</label>
						<input type="text" name="Username" maxlength="60" id="Username" class="ipt_tx" style="width:149px;" tabindex="1" />
						
					</p>
					<p>
						<label>��&nbsp;&nbsp;�룺</label>
						<input type="password" name="Password" maxlength="60" id="Password" class="ipt_tx" style="width:149px;" tabindex="2" autocomplete="off"/>
					</p>
					<%if KS.Setting(34)="1" then%>
					<p>
						<label>��֤�룺</label>
						<input type="text" maxlength="6" name="Verifycode" id="Verifycode" onFocus="this.value='';check.getCode()" class="ipt_tx" style="width:55px;" tabindex="3" autocomplete="off"/>
						<span id="showVerify">������������ʾ</span>
						
					</p>
					<%end if%>
					<p><input type="hidden" name="u1" id="u1"/>
						<input type="submit" tabindex="5"  onClick="return check.CheckForm();" class="btn bnormal" value="��  ¼">
					<a tabindex="6" href="../GetPassword.asp" target="_blank">�������룿</a>  <a tabindex="7" href="../ActiveCode.asp" target='_blank'>�ط�������</a></p>
				</div>
			</form>
			<h2>��û�л�Աͨ��֤�ʺţ�</h2>
			<div class="form_detail">
				<p>
					<input type="button" tabindex="8" id="btn_regist" class="btn bnormal"  onclick="location.href='../../?do=reg'" value="���ھ�ע��" />
				</p>
			</div>
			<h2>���ע�᱾վ��Ա�����軨����30���ӣ�</h2>

            </div>
           
		
        </div>
        <div class="clear"></div>
    </div>

  <br />

	<div id="ft">
		<p>�����п�����Ϣ�������޹�˾ &copy; ��Ȩ����</p>
		<p>��ַ��http://www.kesion.com QQ��9537636 41904294</p>
	</div>
</body>
</html>
        <%
  End Sub
End Class
%> 
