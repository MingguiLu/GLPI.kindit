<!--#include file="../conn.asp"-->
<!--#Include file="../ks_cls/kesion.commoncls.asp"-->
<%
Dim KS:Set KS=New PublicCls
 Dim RS,Email,classid,activecode,ClassInfo,mailid,action
 Dim CheckUrl,MailBodyStr,ReturnInfo
 action=KS.S("Action")
 Email=KS.S("Email")
 activecode=KS.DelSQL(KS.S("ActiveCode"))
 mailid=KS.ChkClng(KS.S("id"))
if action="del" then
   if mailid=0 then ks.die "error!"
   set rs=server.CreateObject("adodb.recordset")
   rs.open "select top 1 * from ks_usermail where id=" & mailid & " and activecode='" & activecode &"'",conn,1,1
   if rs.eof and rs.bof then
     rs.close:set rs=nothing
      KS.Die "<script>alert('�Բ���ɾ����֤��ͨ��,ϵͳ���ܲ����ڴ����䣡');window.close();</script>"
   end if
   email=rs("email")
   rs.close : set rs=nothing
   conn.execute("update ks_usermail set activetf=0 where id=" & mailid)
   'conn.execute("delete from ks_usermail where id=" & mailid)
   KS.Die "<script>alert('��ϲ���ʼ�" & email & "�ڱ�վ�Ķ��ķ�����ȡ����');window.close();</script>"
elseIf Action="cancel" Then
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "SELECT TOP 1 * From KS_UserMail Where Email='" & Email & "'",conn,1,1
  If RS.Eof And RS.Bof Then
    RS.Close :Set RS=Nothing
	KS.AlertHintScript "��������ʼ������ڣ�������������ʼ���ַ���벻Ҫ�Ƿ�������"
  End If
  mailid=rs("id")
  activecode=rs("activecode")
  cassid=rs("classid")
  rs.close : set rs=nothing
  
 IF KS.IsNul(ClassID) Then
		ClassInfo= "ȫ��"
 Else
		   ClassID=Replace(ClassID," ","")
		   ClassIDArr=Split(ClassID,",")
		   For I=0 To Ubound(ClassIDArr)
		     If I<>Ubound(ClassIDArr) Then
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) & "��"
			 Else
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) 
			 End If
		   Next
  End If
  
  CheckUrl = Request.ServerVariables("HTTP_REFERER")
  CheckUrl=KS.GetDomain &"plus/mailsub.asp?action=del&id=" &mailid &"&activecode=" & activecode
  MailBodyStr="<strong>�ڡ�" & KS.Setting(0) & "����վ���ʼ����ķ���ȡ��Ϣȷ�ϣ�</strong><br/>"
  MailBodyStr=MailBodyStr & "ԭ�����Ķ�����Ϣ���£�<br/><br/>"
  MailBodyStr=MailBodyStr & "�������䣺" & email & "<br/>"
  MailBodyStr=MailBodyStr & "���ĵ���Ŀ��<span style='color:blue'>" & ClassInfo & "</span><br/><br/>"
  MailBodyStr=MailBodyStr & "���Ҫȡ�����ģ�������������ɾ�����Ķ��ķ���<br/>"
  MailBodyStr=MailBodyStr & "<a href='" & CheckUrl & "' target='_blank'>" & CheckUrl &"</a><br/><br/>"
  MailBodyStr=MailBodyStr & "<div style='text-align:right'><strong>˵����</strong>���ʼ�ϵͳ�Զ����ͣ�����Ҫ�ظ�!</div>"
  ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "�ʼ����ķ���ȡ��ȷ����", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
  KS.Die "<script>alert('����ȡ�����ķ����������ύ����������ʼ�������Ĳ�����');location.href='../';</script>"
ElseIf action="active" Then
 If mailid=0 Then KS.Die "error!"
 Set RS=Server.CreateObject("adodb.recordset")
 RS.Open "select top 1 * From KS_UserMail Where ActiveCode='" & ActiveCode & "' and id=" & mailid,conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close : Set RS=Nothing
   KS.Die "<script>alert('�������ʧ�ܣ�');window.close();</script>"
 End If
 If RS("ActiveTF")=1 Then
   RS.Close : Set RS=Nothing
   KS.Die "<script>alert('�ö��ķ����Ѽ�����ˣ�����Ҫ�ظ����������');window.close();</script>"
 End If
 RS.Close :Set RS=Nothing
 Conn.Execute("Update KS_UserMail Set ActiveTF=1 Where ID=" & mailid)
   KS.Die "<script>alert('��ϲ�������ʼ����ķ����Ѽ�������᲻���ڵ��յ����ǵĶ����ʼ����񣬸�л����֧�֣�');location.href='../';</script>"
 
ElseIf KS.S("Action")="dosave" Then
 Email=KS.S("Email")
 ClassID=KS.S("ClassID")
 If Not KS.IsValidEmail(Email) Then
    KS.AlertHintScript "�Բ�����������ʼ����Ϸ�!"
 End If
 activecode=KS.MakeRandom(10)
 Set RS=Server.CreateObject("adodb.recordset")
 RS.Open "Select top 1 * From KS_UserMail Where Email='" & Email &"'",conn,1,3
 If RS.Eof Then
   RS.AddNEW
 End If
   RS("ActiveCode")=activecode
   RS("Email")=Email
   RS("ClassID")=ClassID
   If Not KS.IsNul(KS.C("UserName")) Then
   RS("UserName")=KS.C("UserName")
   RS("IsUser")=1
   Else
   RS("IsUser")=0
   End If
   RS("AddDate")=Now
   RS("ActiveTF")=0
   RS.Update
	RS.Close
	Set RS=Nothing
   mailid=KS.ChkClng(Conn.Execute("Select top 1 id From KS_UserMail Where Email='" & Email &"'")(0))
 
 
 IF KS.IsNul(ClassID) Then
		ClassInfo= "ȫ��"
 Else
		   ClassID=Replace(ClassID," ","")
		   Dim ClassIDArr:ClassIDArr=Split(ClassID,",")
		   For I=0 To Ubound(ClassIDArr)
		     If I<>Ubound(ClassIDArr) Then
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) & "��"
			 Else
		      ClassInfo=ClassInfo & KS.C_C(ClassIDArr(i),1) 
			 End If
		   Next
  End If
 
 
 
  CheckUrl = Request.ServerVariables("HTTP_REFERER")
  CheckUrl=KS.GetDomain &"plus/mailsub.asp?action=active&id=" &mailid &"&activecode=" & activecode
  MailBodyStr="<strong>��ȷ�����ڡ�" & KS.Setting(0) & "����վ���ʼ����ķ���</strong><br/>"
  MailBodyStr=MailBodyStr & "�������Ķ�����Ϣ���£�<br/><br/>"
  MailBodyStr=MailBodyStr & "�������䣺" & email & "<br/>"
  MailBodyStr=MailBodyStr & "���ĵ���Ŀ��<span style='color:blue'>" & ClassInfo & "</span><br/><br/>"
  MailBodyStr=MailBodyStr & "�����������Ӽ������Ķ�������<br/>"
  MailBodyStr=MailBodyStr & "<a href='" & CheckUrl & "' target='_blank'>" & CheckUrl &"</a><br/><br/>"
  MailBodyStr=MailBodyStr & "<div style='text-align:right'><strong>˵����</strong>���ʼ�ϵͳ�Զ����ͣ�����Ҫ�ظ�!</div>"
  
  ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "�ʼ����ķ��񼤻��ʼ�", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))

 
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=KS.Setting(1)%> - �ʼ����ķ���</title>
<style>
body{padding:0px;margin:0px;font-size:12px;text-align:center}
#warp{width:960px;margin: 0 auto; border:1px solid #ccc;}
#cphead2010{ background:#000; height: 32px; overflow: hidden; color: #c7e0ff; }
#cphead2010 a{ color: #c7e0ff; font-size: 12px; text-decoration: none; }
#cphead2010 a:hover{ text-decoration: underline; }
#cphead2010 *{ padding: 0; margin: 0; font-size: 12px; }
#cphead2010 img{ border: 0; }
#cphead2010 .cpnav{ height: 32px; }
#cphead2010 .cpnav dd{ float: left; padding: 8px 15px 0 15px; line-height: 20px; }
#cphead2010 .cpnav dd.load{ float: right; }
.box{}
.boxleft{width:610px;float:left;border-right:1px solid #cccccc;}
.boxleft .box1{font-size:14px;color:#fff;
word-spacing:8px;letter-spacing: 4px;font-weight:bold;padding-top:60px;height:70px;background:url(images/box1.gif) no-repeat;}
.boxleft .box2{text-align:left;padding-left:30px;background:url(images/box2.gif) repeat-y;}
.boxleft .box3{height:70px;background:url(images/box3.gif) no-repeat;}
.boxleft .email{width:208px;height:31px;line-height:31px;background:url(images/email.gif) no-repeat;border:0px;padding-left:5px;}
.boxright{text-align:left;width:300px;float:right;padding:20px}
</style>

</head>

<body>
<div id="warp">

<div id="cphead2010">
	<dl class="cpnav">
	<dd><a href="/" target="_blank">��ҳ</a> - <a href="../ask/" target="_blank">�ʴ�</a> - <a href="../club/" target="_blank">��̳</a> - <a href="../user" target="_blank">��Ա</a> - <a href="../space" target="_blank">����</a></dd>
	<dd class="load"><a href="../user/login">��¼</a><span>|</span><a href="../?do=reg" target="_blank">ע��</a></dd>
	</dl>
</div>

<form name="myform" action="mailsub.asp" method="post" />
<div class="box">
	<div class="boxleft">
	  <div class="box1">��ӭʹ�ñ�վ�ʼ����ķ���!</div>
	  <div class="box2">
	  
	  <%If KS.S("Action")="dosave" Then%>
	   <img src='../user/images/regok.jpg' align='left' style="margin:20px"/> <strong>��ϲ�����Ķ����Ѵ�����</strong><br/><br/>
		�����ʼ���<span style='color:red'><%=KS.CheckXSS(Email)%></span><br/>
		���ĵ���Ŀ��<%=ClassInfo%>
		<br/><br/>
		��ע����ȡ����ȷ���ʼ�������Ҫ����ȷ���ʼ��е����ӣ�ȷ����������<br/>�����εĶ��Ĳ����Ż���Ч��
	  <%Else%>
		  <input type="hidden" name="action" id="action" value="dosave" />
		  <img src="images/img13.gif" align="absmiddle"/>
		  <input type="text" name="email" id="email" class="email" value="<%=request("email")%>" maxlength="28"/>
		  <input type="submit" value="ȡ������" onclick="if(document.myform.email.value==''){alert('�����������ʼ�!');return false;}document.getElementById('action').value='cancel';"/>
		  <table border="0" width="100%" align="center">
		   <tr>
			<td colspan="10" height="40" align="left"><strong>��ѡ��������Ȥ����Ŀ�������ѡ���Զ���������Ŀ����ѡ������Ϣ���͸�����</strong></td>
		   </tr>
		  <%
		  KS.LoadClassConfig()
		  Dim Node,I
		  I=0
		  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1 and @ks21=1]")
		   If I=0 Then
			KS.Echo "<tr>"
		   ElseIf I Mod 5=0 Then
			KS.Echo "</tr><tr>"
		   End If
		   KS.Echo "<td><label><input type='checkbox' name='classid' value='" & Node.SelectSingleNode("@ks0").text &"'/>" & Node.SelectSingleNode("@ks1").text & "</label></td>"
		   I=I+1
		  Next
		  %>
		   </tr>
		  </table>
		  <br/>
		  <input type="image" onclick="return(CheckForm())" src="images/img05.jpg"/>
		<%End If%>
	  </div>
	  <div class="box3"></div>
	</div>
	  </form>
	
	<div class="boxright">
	  <p><strong>Ϊʲô�ղ��������ʼ���</strong></p>
		<p>���ܵ�ԭ��<br />
		  1.û�м���ģ����������Ժ���Ҫ����72Сʱ�ڵ������м�����ܽ��յ����š�<br />
		  2.��������,�ʼ����˵�ԭ��,������������ղ����ʼ����ģ�����������ٴζ��ġ�</p>
		<p><strong>���ȡ���ʼ����ģ�</strong></p>
		<p>�����¼��ַ�ʽ��ѡ��<br />
		  1.���󷽵ı���������������ĵ����ݺ��ʼ���ַ������ȡ�����ġ����ɡ� <br />
		  2.�������յ����ʼ��·��С�ȡ���˶��ġ����ӣ�ֱ�ӵ�������ӿ�ȡ�����ġ�</p>
		<p><strong>������¶���ϲ������Ŀ��</strong></p>
		<p>���������������Ķ����ʼ�����ѡ���Լ�ϲ������Ŀ�����ύ���ģ�Ȼ��������ʼ��㼤���������¼���ɡ�</p>
		  
		  
		<p><strong>�ʼ������Ƿ��շѣ�</strong></p>
		<p>���ķ�������ѵģ��������ɱ�վΪ���ṩ�� </p>
	</div>
</div>


</div>
<div style="clear:both;color:#333333;padding:16px;">
 <%=KS.Setting(18)%>
</div>
<script>

function CheckForm()
{
	var form = document.myform;
	var email = form.email.value;
	if (email == '') {
		alert('����д�������䣡');
		form.email.focus();
		return false;
	}
	if (checkMail(email) == false) {
		return false;
	}
	return true;
}

// ����ʼ���ַ�Ƿ�����Ƿ��ַ�
function checkMail(email) {
	var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
	var info = '��Ǹ���ʼ���ַֻ����Ӣ����ĸa��z(�����ִ�Сд)' +
				'������0��9���»���_������-����.��ɣ�' +
				'�����к��ּ����š����ںš�С�ںŵ������ַ���' +
				'����abc@hotmail.com��������������ʼ���ַ��';
	if (!filter.test(email)) {
		alert(info);
		return false;
	}
	return true;
}

</script>
</body>
</html>
