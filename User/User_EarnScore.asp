<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_LogScore
KSCls.Kesion()
Set KSCls = Nothing

Class User_LogScore
        Private KS,KSUser
		Private CurrentPage,totalPut,TotalPages,SQL
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
			MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
		Public Sub loadMain()	
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		Call KSUser.InnerLocation("我要赚积分")
		
	  %>

	 <br/>
	 <style type="text/css">
	  .splittd{height:100px;font-size:16px;padding-left:10px;}
	  .red{color:red;}
	 </style>
	  <script type="text/javascript">
						  function copyToClipboard(txt) {
							 if(window.clipboardData) {
									 window.clipboardData.clearData();
									 window.clipboardData.setData("Text", txt);
							 } else if(navigator.userAgent.indexOf("Opera") != -1) {
								  window.location = txt;
							 } else if (window.netscape) {
								  try {
									   netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
								  } catch (e) {
									   alert("被浏览器拒绝！\n请在浏览器地址栏输入'about:config'并回车\n然后将'signed.applets.codebase_principal_support'设置为'true'");
								  }
								  var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
								  if (!clip)
									   return;
								  var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
								  if (!trans)
									   return;
								  trans.addDataFlavor('text/unicode');
								  var str = new Object();
								  var len = new Object();
								  var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
								  var copytext = txt;
								  str.data = copytext;
								  trans.setTransferData("text/unicode",str,copytext.length*2);
								  var clipid = Components.interfaces.nsIClipboard;
								  if (!clip)
									   return false;
								  clip.setData(trans,null,clipid.kGlobalClipboard);
							 }
								  alert("复制成功！")
						}
		 </script>
	<table border="0" align="center" style="width:100%;">
        <%if KS.Setting(140)="1" Then%>
				<tr>
				  <td class="splittd">
				      <table>
					    <tr>
						  <td><strong>任务名称：</strong></td>
						  <td><span class="red">将本站推荐给朋友将获得积分</span></td>
						</tr>
						<tr>
						 <td><strong>任务介绍：</strong></td>
						 <td>成功推荐一个访问者,您就可以增加 <font color=red><%=KS.Setting(141)%></font> 个积分。赶快行动吧！</td>
						</tr>
						<tr>
						 <td valign="top"><strong>复制代码：</strong></td>
						 <td>
						  <div id="copytext" style="border:1px solid #cccccc;height:45px;width:400px;overflow:scroll"><%=Replace(Replace(KS.Setting(142),"{$UID}",KSUser.UserName),"{$GetSiteUrl}",KS.GetDomain)%></div>
						  <br/><button class="pn" type="button" onClick="copyToClipboard(document.getElementById('copytext').innerHTML);"><strong>复制代码</strong></button>
						 </td>
						 </tr>
						 </table>
													
					</td>
				 </tr>
		 <%end if%>
	  <%if KS.Setting(143)="1" Then%>
		   <tr>
				<td class="splittd"><br/>
				  <table>
				   <tr>
				    <td><strong>任务名称：</strong></td>
					<td><span class="red">引导朋友注册将获得积分</span></td>
				   </tr>
				   <tr>
				    <td><strong>任务介绍：</strong></td>
					<td>成功推荐一个用户注册,您就可以增加 <font color=red><%=KS.Setting(144)%></font> 个积分,同一天内推荐同一个IP的用户注册，只计一次分！</td>
				   </tr>
				   <tr>
				    <td valign="top"><strong>复制代码： </strong>
					</td>
					<td>
					 <div style="border:1px solid #cccccc;height:45px;width:400px;overflow:scroll" id="copytext1"><%=KS.GetDomain%>?do=reg&amp;uid=<%=KSUser.UserName%></div>
									<br/>
									<button class="pn" name="button2" type="button" onClick="copyToClipboard($('#copytext1').text()+'\n<%=Replace(KS.Setting(145),"'","\'")%>');"><strong>复制链接</strong></button>	</td>
						</tr>
						</table>
					 </td>
				   </tr>
			 <%end if%>
			 
				<tr>
				  <td class="splittd">
				      <table>
					    <tr>
						  <td><strong>任务名称：</strong></td>
						  <td><span class="red">邮件邀请好友注册</span></td>
						</tr>
						<tr>
						 <td><strong>任务介绍：</strong></td>
						 <td>给好友发送邀请邮件，好友通过收到的邮件里的链接成功注册为本站会员，您就可以增加 <font color=red><%=KS.Setting(144)%></font> 个积分，同一天内推荐同一个IP的朋友注册，只计一次分！</td>
						</tr>
						<tr>
						 <td valign="top"></td>
						 <td><button class="pn" type="button" onClick="location.href='User_friend.asp?Action=mail'"><strong>我要参加</strong></button>
						 </td>
						 </tr>
						 </table>
													
					</td>
				 </tr>
			 
			 
		 </table>
				
		  <%
  End Sub
    
  
End Class
%> 
