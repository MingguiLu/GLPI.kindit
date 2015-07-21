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
Set KSCls = New UserList
KSCls.Kesion()
Set KSCls = Nothing

Class UserList
        Private KS,KSUser,LoginTF,TopDir,KSR
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
       Public Sub loadMain()
		Call KSUser.Head()
		Call KSUser.InnerLocation("会员首页")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		TopDir=KSUser.GetUserFolder(ksuser.username)
		%>
	
		<style type="text/css">
		  #main .left{float:left;width:500px;}
		  #main .right{float:right;width:230px; margin-top:10px;}
		  #main .userinfo{padding-top:10px; border:#e5e5e5 1px solid; margin-top:10px;}
		  #main .userinfo h1{ font-size:14px; font-weight:bold; padding-left:5px;}
		  #main .userborder{padding:4px;}
          #main .dt td{height:20px;padding-top:5px}
		  #main .dt a{color:#1F7ABC}
		  #main .dt span{color:#999}
		  
		  .visitor{overflow:hidden;padding:10px;}
		  .visitor li a.b{border:1px solid #ccc;padding:1px}
		  .visitor li{float:left;width:100px;text-align:center;}
		  .visitor li img{width:60px;height:60px}
		  
		  #fenye span{display:none}
		  
		*{ padding: 0; margin: 0; }
		h1,h2,h3,h4,h5,h6{ font-size: 12px; font-weight: normal; }
		.fl { float:left }
		.fri { float:right}
		.fri:hover{color:#666666}
		
		#myData .saysth{position:relative; width: 492px; float: left; margin: 0 0 5px 0; background: url(images/writeblog.gif) no-repeat; height: 87px; overflow: hidden; }
		#myData .emotion{ width: 25px; float: left; margin: 4px 10px 0 16px; display: inline; height: 20px; }
		#myData .saysth div{  float: left; border: none; overflow: hidden; }
		#myData .saysth textarea{ padding-top:6px;border: none; margin: 6px 0 0 0; width: 382px; float: left; color: #ABA9A9; background: #fff; font-size: 12px; height: 68px; overflow: auto;}
		.soundBtn{ background: url(images/soundBtn.gif) no-repeat; font-size: 14px; font-weight: bold; color: #666666; text-align:center; line-height:80px; width:59px; height:80px;margin-top:3px; }
		.soundBtn2{ background: url(images/soundBtn2.gif) no-repeat; font-size: 14px; font-weight: bold; color: #fff; text-align:center; line-height:80px; width:59px; height:80px;margin-top:3px; }
		/* 表情 */
		#EmotionsDiv{ display:none;width: 208px; left:10px;top:10px;position: absolute; border: 1px solid #CDCDCD; background: #fff; z-index: 112;padding:1px; }
		#EmotionsDiv img{  float: left;margin:1px; display: inline;border:0  }
		#EmotionsDiv a{ border: 1px solid #E9E9E9;float: left;margin:1px; display: inline ;}
		#EmotionsDiv a:hover{ border: 1px solid #689ACD }
		 
		.cmttextarea{color:#999;width:98%;height:22px;line-height:22px;border:1px solid #ccc;background:#FBFBFB;overflow:auto}
		</style>
		
		<div id="main">
			 <div class="left">
					<div class="userinfo">
					
					
<h1>
					<img src="images/money.gif" align="absmiddle" /> <strong>我的财富</strong> <a href="user_payonline.asp" target="_self"><img src="images/cz.gif" align="absmiddle" border="0"></a> 
					</h1>
					
					  <div class="userborder" >
					  
					   <table width="100%" border="0" cellspacing="0" cellpadding="0">
									  <tr>
										<td height="25"  nowrap>
										您的计费方式为<%if KSUser.ChargeType=1 Then 
										  Response.Write "<font color='#ff6600'>扣点数</font>"
										  ElseIf KSUser.ChargeType=2 Then
										   Response.Write "有效期,到期时间：" & cdate(KSUser.GetUserInfo("BeginDate"))+KSUser.GetUserInfo("Edays") 
										  Else
										   Response.Write "<font color='#ff6600'>永不过期</font>"
										  End If
										  %>
										
										可用资金<font color="green"><%=formatnumber(KSUser.GetUserInfo("Money"),2,-1)%></font>元  ,<%=KS.Setting(45) & "&nbsp;<font color=green>" & formatnumber(KSUser.GetUserInfo("Point"),0,-1) & "</font>" & KS.Setting(46)%>,积分<font color="green"><%=KSUser.GetUserInfo("Score")%></font>分
			                          <%
									   if KS.ChkClng(KSUser.GetUserInfo("UserCardID"))<>0 then
									      Dim RSCard,ValidUnit,ExpireGroupID,ExpireTips
										  Set RSCard=Conn.Execute("Select top 1 * From KS_UserCard Where ID=" & KSUser.GetUserInfo("UserCardID"))
										  If Not RSCard.Eof Then
											 ValidUnit=RSCard("ValidUnit")
											 ExpireGroupID=RSCard("ExpireGroupID")
											 If ValidUnit=1 Then                      '点券
											   If KSUser.GetUserInfo("Point")<=10 And ExpireGroupID<>0 Then
											    ExpireTips="您的" & KS.Setting(45) & "快使用完毕了"
											   End If
											 ElseIf ValidUnit=2 Then                   '有效天数
											   If KSUser.GetUserInfo("Edays")<=7 And ExpireGroupID<>0 Then
											    ExpireTips="您还有" & KSUser.GetUserInfo("Edays") & "天就过期了"
											   End If 
											 ElseIf ValidUnit=3 Then                  '资金
											   If KSUser.GetUserInfo("Money")<=10 And ExpireGroupID<>0 Then
												 ExpireTips="您的账户资金快使用完毕了"
											   End If
											 End If
										  End If
										  RSCard.Close : Set RSCard=Nothing
										  If ExpireTips<>"" and ExpireGroupID<>0  then
										  response.write "<br/><span style='color:red'>温馨提示：您上一次使用充值卡充值，" & ExpireTips & "，<br/>过期后您将自动转为<font color='blue'>"  & KS.U_G(ExpireGroupID,"groupname") & "</font>，为了更好的服务请尽快充值！</span>"
										  end if
									   end if
									  %>   
									  </tr>
									  
					   </table>
					  </div>
					
					 <div class="clear"></div>
					

				
<script type="text/javascript">
function insertface(Val)
	{ 
	  if (Val!=''){
	  var ubb=document.getElementById("CommentContent");
		var ubbLength=ubb.value.length;
		ubb.focus();
		if(typeof document.selection !="undefined")
		{
			document.selection.createRange().text=Val;  
		}
		else
		{
			ubb.value=ubb.value.substr(0,ubb.selectionStart)+Val+ubb.value.substring(ubb.selectionStart,ubbLength);
		}
     }
  }
function sayselect()
{
   document.getElementById('fbHref').className='soundBtn2 fri';
   var txt_obj=document.getElementById("CommentContent");
   if(txt_obj!=null)
   {
       var txt_Value= document.getElementById("CommentContent").value;
       if(txt_Value=="随便说点什么，让好友们知道你的心情、在做什么……（最少10个字）")
       {
           txt_obj.value="";
       }
       
   }
}
function postsay(){
 var c=$("#CommentContent").val();
 if (c==''||c=='随便说点什么，让好友们知道你的心情、在做什么……（最少10个字）'){
  alert('请随便说点什么哦^_^!');
  return false;
 }
 if (c.length<10){
  alert('多写几个字吧^_^！');
  return false;
 }
 $.post("UserAjax.asp",{action:'TalkSave',Content:escape(c)},function(d){
   if (d=="success"){
   alert('成功分享！');
   location.href='index.asp';
   }else{alert(unescape(d));}
 });
}
function showcmt(id){
 $("#sc"+id).toggle();
}
function ThisFocus(id){
 if ($("#c"+id).val()=='我也说一句...'){
  $("#c"+id).val('');
 }
$("#c"+id).attr("style","height:60px;border:2px solid #FFCF5C");
}
function ThisBlur(id){
 if ($("#c"+id).val()==''){
  $("#c"+id).val('我也说一句...');
 }
  $("#c"+id).attr("style","");
  if ($("#c"+id).val()!='我也说一句...'){
$("#c"+id).attr("style","height:60px;");
  }
}
function postcmt(id){
 if ($("#c"+id).val()==''||$("#c"+id).val()=='我也说一句...'){
  alert('您没有输入哦^_^!');
  return false;
 }
 return true;
}
</script>
           
		   
		   
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
						<tr>
						  <td height="25">&nbsp;<strong>您的空间上限容量为 <font color=red><%=round(KSUser.GetUserInfo("SpaceSize")/1024,2)%>M</font></strong> <span id="Sms_txt"></span> &nbsp;<a class="userview" href="User_Files.asp">查看</a></td>
						  <td style="display:none"><img src="images/bar.gif" width="0" height="16" id="Sms_bar" align="absmiddle" /></td>
						</tr>
					  </table>
                 <%
                    response.write showtable("Sms_bar","Sms_txt",KS.GetFolderSize(TopDir)/1024,KSUser.GetUserInfo("SpaceSize"))
                   %>			   
		   
                     
					</div>					
					
					
					<div class="clear"></div>
					
					
					<div class="tabs" style="width:500px">
						<ul>
						<li<%if KS.S("F")="" Then KS.Echo " class=""select"""%>><a href="index.asp">新鲜事</a></li>
						<li<%if KS.S("F")="d" Then KS.Echo " class=""select"""%>><a href="?f=d">随便看看</a></li>
						<li<%if KS.S("F")="g" Then KS.Echo " class=""select"""%>><a href="?f=g">个人动态</a></li>
						<li<%if KS.S("F")="f" Then KS.Echo " class=""select"""%>><a href="?f=f">好友动态</a></li>
						</ul>
					</div>
					
					<% If KS.S("F")="" Then
									 %>
									 
									 <h1>&nbsp;&nbsp;<a href="../space/?<%=KSUser.GetUserInfo("userid")%>/fresh" target="_blank" style="font-size:12px;font-weight:normal;color:#ff6600;">说说我的新鲜事</a></h1>
			<dl id="myData">
            <dd>
              
			<form name="myform" action="index.asp?action=saysave" method="post">
            <div class="saysth" id="saysth_div"> 
			
			<div  id="EmotionsDiv" onmouseleave="this.style.display='none';">
				<%
				dim ns,i
				for i=1 to 24
				 if i<10 then NS=Right("0" & i,2) else NS=i
				%>
				 <a href="#" onclick="insertface('[em<%=ns%>]')"><Img src="../editor/ubb/images/smilies/default/<%=ns%>.gif" /></a>
			<%next%>
			</div><a href="javascript:void(0)" class="emotion" title="插入表情" onclick="document.getElementById('EmotionsDiv').style.display='block';if(document.getElementById('CommentContent').value=='随便说点什么，让好友们知道你的心情、在做什么……（最少10个字）')document.getElementById('CommentContent').value='';" ></a>
              <div>
                <textarea id="CommentContent" name="CommentContent" onclick="javascript:sayselect();" >随便说点什么，让好友们知道你的心情、在做什么……（最少10个字）</textarea>
              </div>
              <!-- js请注意，输入文字后，下面的class变为soundBtn2 -->
              <a id="fbHref" style="cursor:pointer;text-decoration:none" onclick="postsay();" class="soundBtn fri">发 布</a> 
            </div> 
			</form>
           </dd>
		   </dl>
									 
									 <%End If%>
					
					<table border="0" width="100%" class="dt">
					<%
					Dim Param,sqlstr,rs,Totalput,currentpage,MaxPerpage,xml,node,userfield
					MaxPerPage=10
					CurrentPage=KS.ChkClng(KS.S("Page"))
					If CurrentPage=0 Then CurrentPage=1
					
					If KS.IsNul(KS.S("F")) Then
					  MaxPerPage=5
					  sqlstr="select top 200 l.*,userface,RealName from ks_bloginfo l inner join ks_user u on l.username=u.username where l.istalk=1 and l.status=0 order by id desc"
					Else
						if KS.S("F")="d" Then 
						 userfield=",userface"
						 Param=" inner join ks_user u on u.username=l.username where l.username<>'" & KSUser.UserName & "'"
						elseif KS.S("F")="f" Then
						 Param=" inner join ks_friend f on l.username=f.friend where f.username='" & KSUser.UserName & "' and f.accepted=1 and f.ShieldDT=1"
						elseif KS.S("F")="g" Then
						 Param=" Where l.username='"& KSUser.UserName & "'"
						End If
						Sqlstr="select top 500 l.*" & userfield & " from ks_userlog l" & param & " order by l.id desc"
				   End If
					'response.write sqlstr
					             Set RS=Server.CreateObject("AdodB.Recordset")
								 RS.open sqlstr,conn,1,1
								 If RS.EOF And RS.BOF Then
								  RS.Close:SET RS=Nothing
								  KS.Echo "<tr><td class='splittd'>没有记录!</td></tr>"
								 Else
									totalPut = rs.recordcount
									If CurrentPage < 1 Then	CurrentPage = 1
								    If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),rs,"row","root")
									RS.Close:SET RS=Nothing
							    End If
								If IsObject(XML) Then
								  Dim NN:NN=0
								  
								  For Each Node In XML.DocumentElement.SelectNodes("row")
								    
									 KS.Echo "<tr><td class='splittd' valign='top'>"
									
									 
									 
									If userfield<>"" or KS.S("F")="" Then
									 dim userfacesrc:userfacesrc=Node.SelectSingleNode("@userface").text
									 if KS.IsNul(userfacesrc) then userfacesrc="../Images/Face/boy.jpg"
									 if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
									 KS.Echo "<div style='float:left;margin:5px 2px 5px 0px' class='faceborder'><img src='" & userfacesrc & "' width='40' height='40' align='left'/></div></td><td class=""splittd"">"
									End If
									If KS.S("F")<>"" Then
									KS.Echo "<img src='../images/user/log/" & Node.SelectSingleNode("@ico").text & ".gif' align='absmiddle'>"
									KS.Echo "<a href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'>" & Node.SelectSingleNode("@username").text & "</a>"
									KS.Echo " " & Replace(Replace(Replace(Node.SelectSingleNode("@note").text,"{$GetSiteUrl}",KS.GetDomain),"<p>",""),"</p>","") & ""
									Else
									Dim UserName:UserName= Node.SelectSingleNode("@realname").text : If KS.IsNul(UserName) Then UserName= Node.SelectSingleNode("@username").text
									 KS.Echo "<a href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'>" & UserName & "</a> "
									 KS.Echo KSR.ReplaceEmot(Node.SelectSingleNode("@content").text) & ""
									End If
									
									KS.Echo " <span>" 
									If DateDiff("h",Node.SelectSingleNode("@adddate").text,now)>=12 Then
									KS.echo Node.SelectSingleNode("@adddate").text
									Else
									KS.Echo KS.GetTimeFormat(Node.SelectSingleNode("@adddate").text)
									End If
									If KS.S("F")="" Then
									 Dim CmtNum:CmtNum=KS.ChkClng(Node.SelectSingleNode("@totalput").text)
									 Dim CmtNumStr:CmtNumStr="(" & CmtNum & ")"
									 If CmtNum>0 Then CmtNumStr="(<span style='color:red'>" & CmtNum & "</span>)"
									 KS.Echo " <a href='../space/morefresh.asp' target='_blank'>新鲜事</a> <a href=""javascript:void(0)"" onclick=""showcmt(" & Node.SelectSingleNode("@id").text & ")""> 评论" & CmtNumStr & "</a>"
									 KS.Echo "<div id=""sc" & Node.SelectSingleNode("@id").text & """ style="""
									 If NN>0 Then KS.Echo "display:none;"
									 KS.Echo "padding:5px;margin-bottom:6px;margin-left:2px;width:400px;border:1px solid #C1DEFB;background:#E8EFF9;"">"
									 
									 If CmtNum>0 Then
									   Dim RSC:Set RSC=Conn.Execute("Select Top 3 C.AnounName,C.UserName,C.Content,C.Replay,C.replaydate,C.AddDate,U.UserFace,U.UserID,U.RealName From KS_BlogComment C Left Join KS_User U On C.AnounName=u.UserName Where C.LogID=" & KS.ChkClng(Node.SelectSingleNode("@id").text))
									   If Not RSC.Eof Then 
									       Dim UserStr,UserID,Urls,Facestr
										   KS.Echo "<table width='100%' cellspacing='0' cellpadding='0'>"
										   KS.Echo "<tr><td class='splittd' colspan='2'>此条新鲜事共有 <span style='color:red'>" & CmtNum & "</span> 条评论，<a href='../space/?" & Node.SelectSingleNode("@userid").text & "/log/" & KS.ChkClng(Node.SelectSingleNode("@id").text) & "' target='_blank'>查看全部...</a></td></tr>"
										   Do While Not RSC.Eof
										    UserStr=RSC("AnounName")
											If KS.IsNul(UserStr) Then UserStr=RSC("UserName")
											UserID=KS.ChkClng(RSC("UserID"))
											Dim RealName:RealName=RSC("RealName") : If KS.IsNul(RealName) Then RealName=RSC("AnounName")
											Facestr=RSC("UserFace") : If KS.IsNul(Facestr) Then Facestr="images/face/boy.jpg"
											 if left(Facestr,1)<>"/" and lcase(left(Facestr,4))<>"http" then Facestr="../" & Facestr
											If UserID=0 Then Urls="#" Else Urls="" & KS.GetSpaceUrl(UserID)
										    KS.Echo "<tr><td valign='top' style='width:48px;text-align:center'><div style='margin:5px 2px 5px 0px' class='faceborder'><img src='" &facestr & "' width='40' height='40'/></div></td><td class='splittd' style='width:310px'><a href='" & Urls & "' target='_blank'>" & RealName & "</a> " & KS.LoseHtml(RSC("Content")) & " " & KS.GetTimeFormat(RSC("Adddate")) 
											 If Not KS.IsNul(RSC("Replay")) Then
											  KS.Echo "<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>主人 <a href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'>" & UserName & "</a> 回复:</b><br>" & RSc("Replay") & "<br><div align=right>时间:" & rsc("replaydate") &"</div></div>"
											 End If

											KS.Echo "</td></tr>"
										   RSC.MoveNext
										   Loop
										   KS.Echo "</table>"
									   End If
										   RSC.Close : Set RSC=Nothing
									 End If
									 
									 
									 KS.Echo "<form  name=""form" & Node.SelectSingleNode("@id").text & """ action=""../space/writecomment.asp"" method=""post""><input type=""hidden"" name=""action"" value=""CommentSave""/><input type=""hidden"" name=""id"" value=""" & Node.SelectSingleNode("@id").text & """/><input type=""hidden"" name=""AnounName"" value=""" & KSUser.UserName & """/><input type=""hidden"" name=""from"" value=""1""/>"
									 KS.Echo "<textarea name=""Content"" id=""c" & Node.SelectSingleNode("@id").text & """ class=""cmttextarea"" onblur=""ThisBlur(" & Node.SelectSingleNode("@id").text & ")"" onfocus=""ThisFocus(" & Node.SelectSingleNode("@id").text & ")"" cols=""50"" rows=""2"">我也说一句...</textarea><br/><div style=""margin:4px 0px 4px 0px""><input type=""submit"" class=""button"" onclick=""return(postcmt(" & Node.SelectSingleNode("@id").text & "))"" value=""发表""/></div></form></div>"

									End If
									KS.Echo "</span>"
									
									KS.Echo "</td></tr>"
									NN=NN+1
								  Next
								End If
					           XML=Empty : Set Node=Nothing
					%>
					    
						</table><%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
					           <div class='splittd clear'>
								  </div>
					
					
					
			</div>
			<div class="right">
			   
			   <div class="tbox">
		        <div class="t_head">活动公告</div>
				 <%
				  Dim KSFObj:Set KSFobj=New refreshFunction
				  KS.Echo KSFObj.getLabel("{Tag:GetAnnounceList labelid=""0"" announcetype=""2"" owidth=""450"" oheight=""400"" width=""225"" height=""100"" speed=""1"" showstyle=""1"" opentype=""1"" listnumber=""10"" titlelen=""30"" showauthor=""0"" contentlen=""100"" navtype=""1"" titlecss="""" channelid=""9990"" ajaxout=""false""}{/Tag}")
				  Set KSFobj=Nothing
				 %>
				
				  
			   </div>
			   <div class="tbox" style="margin-top:10px;">
		        <div class="t_head">最近谁来看过我</div>
				<div class="visitor">
				<%
				Dim user_face,Visitors
				Set RS=Conn.Execute("Select top 10 b.sex,a.Visitors,b.userface,a.adddate,b.isonline,b.userid from KS_BlogVisitor a inner join KS_User b on a.Visitors=b.username where a.username='" & KSUser.UserName & "' order by a.adddate desc ,id desc")
				If Not RS.Eof Then Set XML=KS.RsToXml(Rs,"row","")
				RS.Close:Set RS=Nothing
			    If IsObject(XML) Then
				  For Each Node In XML.DocumentElement.SelectNodes("row") 
				    user_face=Node.SelectSingleNode("@userface").text
					Visitors =Node.SelectSingleNode("@visitors").text
					If user_face="" or isnull(user_face) then 
					 if Node.SelectSingleNode("@sex").text="男" then  user_face="images/face/boy.jpg" else user_face="images/face/girl.jpg"
					End If
			        If lcase(left(user_face,4))<>"http" and left(user_face,1)<>"/" then user_face=KS.Setting(2) & "/" & user_face
			         KS.Echo "<li><a class='b' href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'><img src='" & User_face & "' border='0'></a><br/><a class='user_name' href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'>" & Visitors & "</a><br />状态："
					 If Node.SelectSingleNode("@isonline").Text="1" Then KS.Echo "<font color=red>在线</font>" Else KS.Echo "离线"
					 KS.Echo "</li>"
				  Next
				  XML=Empty : Set Node=Nothing
				Else
				    KS.Echo "没有访问记录,要加油哦^_^!"
				End If
				%>
				</div>
			   </div>
			
			</div>
		<div class="clear"></div>
		
		 
		 
		  		</div>

		<%
  End Sub
  
	   '（图片对象名称，标题对象名称，更新数，总数）
		Function ShowTable(SrcName,TxtName,str,c)
		Dim Tempstr,Src_js,Txt_js,TempPercent,SrcWidth
		If C = 0 Then C = 99999999
		Tempstr = str/C
		TempPercent = FormatPercent(tempstr,0,-1)
		Src_js = "document.getElementById(""" + SrcName + """)"
		Txt_js = "document.getElementById(""" + TxtName + """)"
			ShowTable = VbCrLf + "<script>"
			SrcWidth=FormatNumber(tempstr*300,0,-1) : If SrcWidth>500 Then SrcWidth="100%"
			ShowTable = ShowTable + Src_js + ".width=""" & SrcWidth & """;"
			ShowTable = ShowTable + Src_js + ".title=""容量上限为："&c/1024&" MB，已用（"&FormatNumber(str/1024,2)&"）MB！"";"
			ShowTable = ShowTable + Txt_js + ".innerHTML="""
			If FormatNumber(tempstr*100,0,-1) < 80 Then
				ShowTable = ShowTable + "已使用:" & TempPercent & """;"
			ElseIf FormatNumber(tempstr*100,0,-1)>100 Then
				ShowTable = ShowTable + "<font color=\""red\"">可用空间已使用完毕,请赶快清理！</font>"";"
			Else
				ShowTable = ShowTable + "<font color=\""red\"">已使用:" & TempPercent & ",请赶快清理！</font>"";"
			End If
			ShowTable = ShowTable + "</script>"
		End Function
End Class
%> 
