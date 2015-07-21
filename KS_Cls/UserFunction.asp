<%
Sub Echo(sStr)
	 Response.Write sStr 
	 'Response.Flush()
End Sub
  
public Sub Scan(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "[#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = Parse(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

Function Parse(sTemplate, iPosBegin)
	Dim iPosCur, sToken, sTemp,MyNode
	iPosCur      = InStr(iPosBegin, sTemplate, "]")
	sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
	iPosBegin    = iPosCur+1
	select case Lcase(sTemp)
		case "pubtips"  
		  If Action="Edit" Then
		    echo "修改" & KS.C_S(Channelid,3)
		  Else
		    echo "发布" & KS.C_S(Channelid,3)
		  End If
		case "selectclassid"
		   Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) 
		case "status"
		   if action="Edit" Then
		     If RS("Verific")<>1 Then
			  if rs("verific")=2 Then
		       echo "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' checked>放入草稿</label>"
			  else
		       echo "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' value='2'>放入草稿</label>"
			  end if
			 Else
			  echo "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
		   Else
		    echo "<label style='color:#999999'><input type='checkbox' name='status' value='2'>放入草稿</label>"
		   End If
		case "content"
		  If Action="Edit" Then
		   if KS.C_S(ChannelID,6)=1 Then if not KS.IsNul(rs("ArticleContent")) then echo Server.HtmlEncode(rs("ArticleContent"))
		  End If
		case else
		   on error resume next
		   Dim II,DV
		   if instr(sTemp,"|select")<>0 then  '下拉及联动
			 For II=0 To Ubound(UserDefineFieldArr,2)
			   If Lcase(UserDefineFieldArr(0,Ii))=Lcase(split(sTemp,"|")(0)) Then
			      If Action="Edit" Then DV=RS(Trim(UserDefineFieldArr(0,ii))) Else DV=UserDefineFieldArr(4,ii)
			       KS.Echo  KSUser.GetSelectOption(ChannelID,UserDefineFieldValueStr,UserDefineFieldArr,UserDefineFieldArr(3,II),UserDefineFieldArr(0,ii),UserDefineFieldArr(7,ii),UserDefineFieldArr(5,ii),DV)
				 Exit For
			   End If
			 Next
		   elseif instr(sTemp,"|radio")<>0 then  '单选
			 For II=0 To Ubound(UserDefineFieldArr,2)
			   If Lcase(UserDefineFieldArr(0,Ii))=Lcase(split(sTemp,"|")(0)) Then
			      If Action="Edit" Then DV=RS(Trim(UserDefineFieldArr(0,ii))) Else DV=UserDefineFieldArr(4,ii)
			       KS.Echo  KSUser.GetRadioOption(UserDefineFieldArr(0,ii),UserDefineFieldArr(5,ii),DV)
				 Exit For
			   End If
			 Next
		   elseif instr(sTemp,"|checkbox")<>0 then  '多选
			 For II=0 To Ubound(UserDefineFieldArr,2)
			   If Lcase(UserDefineFieldArr(0,Ii))=Lcase(split(sTemp,"|")(0)) Then
			      If Action="Edit" Then DV=RS(Trim(UserDefineFieldArr(0,ii))) Else DV=UserDefineFieldArr(4,ii)
			       KS.Echo  KSUser.GetCheckBoxOption(UserDefineFieldArr(0,ii),UserDefineFieldArr(5,ii),DV)
				 Exit For
			   End If
			 Next
		   elseif instr(sTemp,"|unit")<>0 then  '单位
			 For II=0 To Ubound(UserDefineFieldArr,2)
			   If Lcase(UserDefineFieldArr(0,Ii))=Lcase(split(sTemp,"|")(0)) Then
			       If Action="Edit" Then DV=RS(Trim(UserDefineFieldArr(0,ii))&"_unit") Else DV=UserDefineFieldArr(4,ii)
			       KS.Echo  KSUser.GetUnitOption(UserDefineFieldArr(0,ii),UserDefineFieldArr(12,ii),DV)
				 Exit For
			   End If
			 Next
		   elseif action="Edit" Then
		     echo rs(trim(stemp))
		   Elseif left(lcase(sTemp),3)="ks_" then
		     echo server.htmlencode(GetDiyFieldValue(UserDefineFieldArr,sTemp))
		   End If
		   if err.number<>0 then ks.die "<script>alert('字段没有找到,请检查后台模型管理->修改->投稿选项的录入模板!');</script>" : err.clear
	end select
	Parse    = iPosBegin
 End Function
 
 
'=========================扫描会员中心主体框架 增加于2010年6月========================================

Public Sub Kesion()
         Dim LoginTF:LoginTF=Cbool(KSUser.UserLoginChecked)

        '==========================设置在线状态================================
        If Request.QueryString("action")="offline" then
		 session("setonlinestatus")="true"
		 Conn.Execute("Update KS_User Set isonline=0 where username='" & KSUser.UserName &"'")
		 If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@isonline").Text=0
		 Response.Redirect Request.ServerVariables("HTTP_REFERER")
		ElseIf Request.QueryString("action")="setonline" Then
		 session("setonlinestatus")="true"
		 Conn.Execute("Update KS_User Set isonline=1 where username='" & KSUser.UserName &"'")
		 If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@isonline").Text=1
		 Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End If
		'===================================================================
		
		 Dim FileContent,MainUrl,RequestItem,TemplateFile
		 Dim KSR,ParaList
		 FCls.RefreshType = "Member"   '设置当前位置为会员中心
		 Set KSR = New Refresh
		 TemplateFile=KS.Setting(116)
		 If LoginTF=True Then  TemplateFile=KS.U_G(KSUser.GroupID,"templatefile")
		 If trim(TemplateFile)="" Then TemplateFile=KS.Setting(116)
         If trim(TemplateFile)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		 FileContent = KSR.LoadTemplate(TemplateFile)
		 If Trim(FileContent) = "" Then FileContent = "模板不存在!"
		  FileContent = KSR.KSLabelReplaceAll(FileContent)
		 Set KSR = Nothing
		 ScanTemplate FileContent
End Sub	
 
public Sub ScanTemplate(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "{#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = ParseTemplate(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

Function ParseTemplate(sTemplate, iPosBegin)
		Dim iPosCur, sToken, sTemp,MyNode
		iPosCur      = InStr(iPosBegin, sTemplate, "}")
		sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
		iPosBegin    = iPosCur+1
		select case Lcase(sTemp)
			case "showusermain"  loadMain
			case "showuserinfo"  GetUserBasicInfo
			case "showtoptips"  TopTips
			case "showmymenu"  ShowMyMenu
		end select
		 ParseTemplate=iPosBegin
End Function

Sub GetUserBasicInfo()
		%>
		<div class="mem_left_top">
			<div class="mem_left_photo">
			<%
			IF KS.IsNul(KS.C("UserName")) Then
			   KS.Echo "<img src=""../images/face/boy.jpg"" width=""62"" height=""58"" alt=""个人形象"" />"
			Else
			  Dim UserFaceSrc:UserFaceSrc=KSUser.GetUserInfo("UserFace")
			  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
			  KS.Echo ("<img src=""" & UserFaceSrc & """ alt=""" & KSUser.GetUserInfo("RealName") & """ border=""1"" width=""62"" height=""58"">")
			End If
			%>
			</div>
			<div class="mem_left_name">
			 <ul>
				<%
				 KS.Echo ("<li>昵称：" & KSUser.UserName & "</li><li>组别：<span style=""curson:pointer"" title=""" & KS.U_G(KSUser.GroupID,"groupname") & """>" & KS.Gottopic(KS.U_G(KSUser.GroupID,"groupname"),12) & "</span></li>")
				 %>
				 <li>状态：<span class="rl" style="cursor:pointer" onMouseover="$('#userstatus').show();"><%if KSUser.GetUserInfo("IsOnline")="1" then response.write "<img src=""images/online.gif"" align='absmiddle'> 在线" else response.write "<img src=""images/notonline.gif"" align='absmiddle'> 隐身"%>
										<div id="userstatus" class='abs' style="top:15px;width:100px" onMouseOut="$('#userstatus').hide();">
											<dl><img src="images/downline.png" align="absmiddle" />  <a href="?action=offline">隐身离线</a></dl>
											<dl><img src="images/online.png" align="absmiddle" /> <a href="?action=setonline">我在线上</a></dl>
										</div>
										</span>
				 </li>
				 <li><a href="User_EditInfo.asp?Action=face" style="color:#0066CC;text-decoration:underline">修改头像</a> <a href="user_editinfo.asp" style="color:#0066CC;text-decoration:underline">用户资料</a></li>
			 </ul>
		  </div>
		</div>
		<%
End Sub
Sub TopTips()
		    Dim Str
			If KSUser.GetUserInfo("realname")<>"" then
		 	KS.Echo ("<strong><span style='color:green'>"  & KSUser.GetUserInfo("realname") & "")
			Else
		 	KS.Echo ("<strong><span style='color:green'>" & KSUser.UserName)
			End If
			if KSUser.GetUserInfo("Sex")="男" then
			 KS.Echo "先生</span></strong> "
			Else
			 KS.Echo "女士</span></strong> "
			End If
			If (Hour(Now) < 6) Then
            Str= "凌晨好，"
			ElseIf (Hour(Now) < 9) Then
			Str= "早上好，"
			ElseIf (Hour(Now) < 12) Then
			Str= "上午好，"
			ElseIf (Hour(Now) < 14) Then
			Str= "中午好，"
			ElseIf (Hour(Now) < 17) Then
			Str= "下午好，"
			ElseIf (Hour(Now) < 18) Then
			Str= "傍晚好，"
			Else
			Str= "晚上好，"
			End If
			KS.Echo str & " 欢迎来到会员中心！<br/>"
			
			
	  If KS.SSetting(0)<>0 Then  '判断有没有开通空间
			 dim spacedomain,predomain
			 If KS.SSetting(14)<>"0" and not conn.execute("select top 1 username from ks_blog where username='" & ksuser.username & "'").eof Then
			   predomain=conn.execute("select top 1 [domain] from ks_blog where username='" & ksuser.username & "'")(0)
			 end if
			 if Not KS.IsNul(predomain) then
				if instr(predomain,".")=0 then
					spacedomain="http://" & predomain & "." & KS.SSetting(16)
				else
				  spacedomain="http://" & predomain
				end if
			 else
					 If KS.SSetting(21)="1" Then
						 spacedomain=KS.GetDomain & "space/" & ks.c("userid")
					 Else
						 spacedomain=KS.GetDomain & "space/?" & ks.c("userid")
					 End If
			 end if
		 If KSUser.CheckPower("s01")=true then
		   KS.Echo ("我的空间地址:<a href='" & spacedomain & "' target='_blank'>" & spacedomain & "</a>&nbsp;&nbsp; <span  class='rl' onMouseOver=""$('#myspace').show();"" onMouseOut=""$('#myspace').hide();"" style=""color:#ff6600;""><span style='font-weight:bold;font-size:14px'>空间设置</span><img src='images/dico.gif' align='absmiddle'/> <div class='abs' style='width:100px;padding-top:-10px;top:16px;' id='myspace'>")
		    If Conn.Execute("Select top 1 ID From ks_EnterPrise Where UserName='" &KSUser.UserName & "'").eof Then
			   KS.Echo ("<dl> <a href='User_Blog.asp?Action=BlogEdit'><img src='images/menu_icon.gif'> 个人空间设置</a></dl>")
			   KS.Echo ("<dl> <a href='user_Enterprise.asp'><img src='images/menu_icon.gif'> 升级为企业空间</a></dl>")
			Else
			   KS.Echo ("<dl> <a href='User_Blog.asp?Action=BlogEdit'><img src='images/menu_icon.gif'> 企业空间设置</a></dl>")
			   KS.Echo ("<dl> <a href='User_Blog.asp?action=Template'><img src='images/menu_icon.gif'> 空间模板绑定</a></dl>")
			End If

			
		   KS.Echo "</div></span>"
		 ELSEIf KSUser.GetUserInfo("usertype")="1" Then
		    KS.Echo "<a href='../company/show.asp?username=" & ksuser.username & "' target='_blank'>" & KS.GetDomain & "company/show.asp?username=" & KSUser.UserName &"</a>"
		 End If
	End If
 End Sub
 
 Sub ShowMyMenu()
   If KS.SSetting(0)=1 Then  '开通空间则退出
		 If KSUser.CheckPower("s02")=true And KSUser.CheckPower("s01")=true then
			 Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			 Response.Write "<img src=""images/icon7.png"" align=""absmiddle"" /> <a href=""User_Blog.asp"">博文</a>"
			 Response.Write "<span><a href=""User_Blog.asp?Action=Add"">+发表</a></span></li>"
		 End If
		 If KSUser.CheckPower("s05")=true And KSUser.CheckPower("s01")=true Then
			 Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			 Response.Write "<img src=""images/icon2.png"" align=""absmiddle"" /> <a href=""User_Photo.asp"">相册</a>"
			 Response.Write "<span><a href=""User_Photo.asp?Action=Add"">+上传</a></span></li>"
		 End If
		 If KSUser.CheckPower("s06")=true And KSUser.CheckPower("s01")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon19.png"" align=""absmiddle"" /> <a href=""User_Team.asp"">圈子</a>"
			Response.Write "<span><a href=""User_Team.asp?action=CreateTeam"">+创建</a></span></li>"
		 End If

	  If Conn.Execute("Select top 1 ID From ks_EnterPrise Where UserName='" &KSUser.UserName & "'").eof Then '个人空间
		 If KSUser.CheckPower("s04")=true And KSUser.CheckPower("s01")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon20.png"" align=""absmiddle"" /> <a href=""User_music.asp"">音乐</a>"
			Response.Write "<span><a href=""User_Music.asp?action=addlink"">+添加</a></span></li>"
		 End If
		 If KSUser.CheckPower("s10")=true Then 
			If KS.C_S(5,21)="1" Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon6.png"" align=""absmiddle"" /> <a href=""user_myshop.asp"">商品</a>"
			Response.Write "<span><a href=""user_myshop.asp?ChannelID=5&Action=Add"">+发布</a></span></li>"
			End If
		 End IF
	  Else   '企业空间
	      if KSUser.CheckPower("s10")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon6.png"" align=""absmiddle"" /> <a title='企业产品管理' href=""user_myshop.asp"">产品</a>"
			Response.Write "<span><a href=""user_myshop.asp?ChannelID=5&Action=Add"">+发布</a></span></li>"
		  End If
		  if KSUser.CheckPower("s11")=true And KSUser.CheckPower("s01")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon21.png"" align=""absmiddle"" /> <a title='企业新闻管理' href=""user_EnterpriseNews.asp"">动态</a>"
			Response.Write "<span><a href=""user_EnterpriseNews.asp?Action=Add"" title='发布企业新闻'>+发布</a></span></li>"
		  end if
		  if KSUser.CheckPower("s12")=true And KSUser.CheckPower("s01")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a title='关键词广告管理管理' href=""user_EnterpriseAD.asp"">广告</a>"
			Response.Write "<span><a href=""user_EnterpriseAD.asp?Action=Add"">+发布</a></span></li>"
		  end if
		  if KSUser.CheckPower("s13")=true And KSUser.CheckPower("s01")=true then
			  	Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
				Response.Write "<img src=""images/icon22.png"" align=""absmiddle"" /> <a title='企业荣誉证书管理' href=""user_Enterprisezs.asp"">荣誉</a>"
				Response.Write "<span><a href=""user_Enterprisezs.asp?Action=Add"">+发布</a></span></li>"
		  End If
	  End If
  End If
   		 If KSUser.CheckPower("s03")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon18.png"" align=""absmiddle"" /> <a href=""User_friend.asp"">好友</a>"
			Response.Write "<span><a href=""User_Friend.asp?action=addF"">+寻找</a></span></li>"
		 End If

  
  
  	 
'模型的投稿
if KSUser.CheckPower("s18")<>false Then 
			 Dim Node,Ico,ItemUrl
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig()
			 
			 For Each Node In Application(KS.SiteSN&"_ChannelConfig").DocumentElement.SelectNodes("channel[@ks21=1 and @ks36!=0 and @ks0!=5 and @ks0!=6]")
				Ico=Node.SelectSingleNode("@ks51").text
				If KS.IsNul(Ico) Then Ico="images/icon7.png"
				Select Case KS.ChkClng(Node.SelectSingleNode("@ks6").text) 
				  Case 1 ItemUrl="User_MyArticle.asp"
				  Case 2 ItemUrl="User_MyPhoto.asp"
				  Case 3 ItemUrl="User_MySoftWare.asp"
				  Case 4 ItemUrl="User_Myflash.asp"
				  Case 5 ItemUrl="User_MyShop.asp"
				  Case 7 ItemUrl="User_MyMovie.asp"
				  Case 8 ItemUrl="User_MySupply.asp"
				  Case 9 ItemUrl="User_MyExam.asp"
			   End Select
			   		Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
					Response.Write "<img src=""" & Ico & """ align=""absmiddle"" /> <a href=""" & ItemUrl &"?channelid="& Node.SelectSingleNode("@ks0").text & """>" & Node.SelectSingleNode("@ks52").text & "</a>"
					If KS.ChkClng(Node.SelectSingleNode("@ks6").text) =9 Then
					Response.Write "<span><a href=""User_MyExam.asp?action=record"">+记录</a></span></li>"
					Else
					Response.Write "<span><a href=""" & ItemUrl &"?channelid="& Node.SelectSingleNode("@ks0").text & "&Action=Add"">+发布</a></span></li>"
					End If
			 Next
	   End If
		 
		 '求职
		If KS.C_S(10,21)=1 Then
			If KSUser.GetUserInfo("UserType")=0 Then
				If KSUser.CheckPower("s14")=true Then 
					 If KS.C_S(10,21)="1" Then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a href=""User_JobResume.asp"">找工作</a>"
						Response.Write "<span><a href=""User_JobResume.asp"">+简历</a></span></li>"

					 End If
				End If
			Else			 
			if KSUser.CheckPower("s14")=true  then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a href=""user_Enterprise.asp?action=job"">找人才</a>"
						Response.Write "<span><a href=""User_JobCompanyZW.asp?Action=Add"">+发布</a></span></li>"
			 end if
            End If
		End If
		
		 If KS.C_S(5,21)=1 Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon9.png"" align=""absmiddle"" /> <a href=""user_order.asp"">订单</a>"
			Response.Write "<span><a href=""user_order.asp?action=coupon"">优惠券</a></span></li>"
		 End if
		
		 If KSUser.CheckPower("s07")=true Then 
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon7.png"" align=""absmiddle"" /> <a href=""User_Class.asp"">专栏</a>"
			Response.Write "<span><a href=""User_Class.asp?Action=Add"">+创建</a></span></li>"
		 End If
		 If KSUser.CheckPower("s09")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon11.png"" align=""absmiddle"" /> <a href=""User_Askquestion.asp"">问答</a>"
			Response.Write "<span><a  href=""../ask/a.asp"" target=""_blank"">+提问</a></span></li>"
		 End If
		 If KSUser.CheckPower("s19")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon11.png"" align=""absmiddle"" /> <a href=""User_mytopic.asp"">论坛</a>"
			Response.Write "<span><a  href=""User_mytopic.asp?action=fav"">收藏帖</a></span></li>"
		 End If
		 If KSUser.CheckPower("s20")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon13.png"" align=""absmiddle"" /> <a href=""User_ItemSign.asp"">签收</a>"
			Response.Write "<span><a  href=""User_ItemSign.asp"">查看</a></span></li>"
		 End If
         if KSUser.CheckPower("s16")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon13.png"" align=""absmiddle"" /> <a href=""User_favorite.asp"">收藏</a>"
			Response.Write "<span><a  href=""User_favorite.asp"">查看</a></span></li>"
		 End If
         if KSUser.CheckPower("s17")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon15.png"" align=""absmiddle"" /> <a href=""user_feedback.asp"">投诉</a>"
			Response.Write "<span><a  href=""user_feedback.asp?Action=Add"">+发布</a></span></li>"
		 End If
End Sub
 
'------扫描会员中心主体框架------

 
 
 
 '取得某个字段的默认值
 Function GetDiyFieldValue(F_Arr,FieldName)
		     Dim I,v
			 For I=0 To Ubound(F_Arr,2)
			     If Lcase(F_Arr(0,I))=Lcase(FieldName) Then
				   v=F_Arr(4,i)
				   Exit For
				 End If
			Next
			If Instr(V,"|")<>0 Then
			 V=LFCls.GetSingleFieldValue("select top 1 " & Split(V,"|")(1) & " from " & Split(V,"|")(0) & " where username='" & KSUser.UserName & "'") 
			End If
			GetDiyFieldValue=v
 End Function

'参数 isTemplate true 后台生成表单模板调用,channelid 模型id, id 编辑时的文章ID
Function GetInputForm(IsTemplate,ChannelID,id,KSUser,RS)
  Dim F_B_Arr,F_V_Arr,UserDefineFieldValueStr
  Dim ClassID,Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,Intro,FullTitle,ReadPoint,Province,City,UserDefineFieldArr,I,SelButton,MapMarker
  F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
  F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
  
if IsObject(RS) And IsTemplate=false Then
	If Not RS.Eof Then
		     If KS.C_S(ChannelID,42) =0 And RS("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
			   RS.Close():Set RS=Nothing
			   KS.ShowTips "error",server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
			   KS.Die ""
			 End If
		     ClassID  = RS("Tid")
			 Title    = RS("Title")
			 KeyWords = RS("KeyWords")
			 Author   = RS("Author")
			 Origin   = RS("Origin")
			 Content  = RS("ArticleContent")
			 Verific  = RS("Verific")
			 If Verific=3 Then Verific=0
			 PhotoUrl   = RS("PhotoUrl")
			 Intro    = RS("Intro")
			 FullTitle= RS("FullTitle")
			 ReadPoint= RS("ReadPoint")
			 Province = RS("Province")
			 City     = RS("City")
			 If F_B_Arr(25)="1" Then	MapMarker=RS("MapMarker")
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
					  Dim UnitOption
					  If UserDefineFieldArr(11,I)="1" Then
					   UnitOption="@" & RS(UserDefineFieldArr(0,I)&"_Unit")
					  Else
					   UnitOption=""
					  End If
				  If i=0 Then
				    UserDefineFieldValueStr=RS(UserDefineFieldArr(0,I)) &UnitOption & "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & RS(UserDefineFieldArr(0,I)) & UnitOption & "||||"
				  End If
				Next
			  End If
		   End If
		   RS.Close:Set RS=Nothing
		   SelButton=KS.C_C(ClassID,1)
		Else
		 If IsTemplate=false Then
		     Call KSUser.CheckMoney(ChannelID)
			 Author=KSUser.GetUserInfo("RealName")
			 Origin=LFCls.GetSingleFieldValue("SELECT top 1 CompanyName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'")
			 ClassID=KS.S("ClassID")
			 If ClassID="" Then ClassID="0"
			 If ClassID="0" Then
			 SelButton="选择栏目..."
			 Else
			 SelButton=KS.C_C(ClassID,1)
			 End If
			 ReadPoint=0 : Verific=0
		 Else
		    Title="[#Title]"
			FullTitle="[#FullTitle]"
			KeyWords="[#KeyWords]"
			Author="[#Author]"
			Origin="[#Origin]"
			Province="[#Province]"
			City="[#City]"
			Author="[#Author]"
			Intro="[#Intro]"
			Content="[#Content]"
			PhotoUrl="[#PhotoUrl]"
			ReadPoint="[#ReadPoint]"
			Verific="[#Verific]"
			MapMarker="[#MapMarker]"
			 UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
			 If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				  If i=0 Then
				    UserDefineFieldValueStr="[#" & UserDefineFieldArr(0,I) & "]" & "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & "[#" & UserDefineFieldArr(0,I) & "]"  & "||||"
				  End If
				Next
			  End If
		 End If
		End If
		%><table  width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
 <tr class="title">
  <td colspan=2 align=center><%
	      If IsTemplate Then
		  Response.Write "[#PubTips]"
		  ElseIF ID<>0 Then
			  response.write "修改" & KS.C_S(ChannelID,3)
		  Else
		      response.write "发布" & KS.C_S(ChannelID,3)
		 End iF%></td>
 </tr>
 <tr class="tdbg">
  <td width="12%"  height="25" align="center"><span><%=F_V_Arr(1)%>：</span></td>
  <td width="88%"><%
				If IsTemplate Then
				  Response.Write "[#SelectClassID]"
				Else
				 Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) 
				End If
			 %></td>
 </tr>
 <tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(0)%>：</span></td>
    <td><input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" /><span style="color: #FF0000">*</span></td>
 </tr>
<%if F_B_Arr(2)=1 Then%> <tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(2)%>：</span></td>
    <td><input class="textbox" name="FullTitle" type="text" style="width:250px; " value="<%=FullTitle%>" maxlength="100" /><span class="msgtips"> 完整标题，可留空</span></td>
 </tr>
<%End If%>
<%if F_B_Arr(5)=1 Then%> <tr class="tdbg">
    <td height="25" align="center"><span><%=F_V_Arr(5)%>：</span></td>
    <td><input name="KeyWords"  class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:220px; " /><a href="javascript:void(0)" onclick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开</span></td>
  </tr>
<%end if%>
<%if F_B_Arr(6)=1 Then%> <tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(6)%>：</span></td>
    <td height="25"><input name="Author" class="textbox" type="text" id="Author" style="width:220px; " value="<%=Author%>" maxlength="30" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的作者<span></td>
  </tr>
<%end if%>
<%if F_B_Arr(7)=1 Then%>  <tr class="tdbg">
   <td  height="25" align="center"><span><%=F_V_Arr(7)%>：</span></td>
   <td><input class="textbox" name="Origin" type="text" id="Origin" style="width:220px; " value="<%=Origin%>" maxlength="100" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的来源<span></td>
  </tr>
<%end if%>
<%if F_B_Arr(23)="1" Then%>	<tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(23)%>：</span></td>
    <td><script src="../plus/area.asp" type="text/javascript"></script>
									  <script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=province%>');
								  <%end if%>
							  <%if City<>"" Then%>
							  $('#City')[0].options[1]=new Option('<%=City%>','<%=City%>');
							  $('#City')[0].options(1).selected=true;
							  <%end if%>
							</script>
	</td>
 </tr>
<%end if%>
<%if F_B_Arr(25)="1" Then%> <tr class="tdbg">
    <td height="25" align="center"><span><%=F_V_Arr(25)%>：</span></td>
    <td>经纬度：<input value="<%=MapMarker%>" type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='images/edit_add.gif' align='absmiddle' border='0'>添加电子地图标志</a>
	 <script type="text/javascript">
		  function addMap(){
		  new KesionPopup().PopupCenterIframe('电子地图标注','../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),760,430,'auto');
		  }
		  </script>
	</td>
  </tr>
<%end if%>
<%Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)%>
<%if F_B_Arr(8)=1 Then%> <tr class="tdbg">
   <td  height="25" align="center"><span><%=F_V_Arr(8)%>：</span><br><input name='AutoIntro' type='checkbox' checked value='1'><font color="#FF0000">自动截取内容的200个字作为导读</font></td>
   <td><textarea class='textarea' name="Intro" style='width:95%;height:95px'><%=intro%></textarea></td>
  </tr>
<%end if%>
<%if F_B_Arr(9)=1 Then%> <tr class="tdbg">
   <td><%=F_V_Arr(9)%>:<br><img src="images/ico.gif" width="17" height="12" /><font color="#FF0000">如果<%=KS.C_S(ChannelID,3)%>较长可以使用分页标签：[NextPage]</font></td>
   <td><%
								If F_B_Arr(21)=1 Then
	%>
			      <table border='0' width='100%' cellspacing='0' cellpadding='0'>
			       <tr><td height='35' width=70>&nbsp;<strong><%=F_V_Arr(21)%>:</strong></td><td width="170"><input onclick='PopInsertAnnex()' class='button' type='button' name='annexbtn' value='插入已存在的附件'/>&nbsp;</td><td><iframe id='UpFileFrame' name='UpFileFrame' src='User_UpFile.asp?Type=File&ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='200' height='24'></iframe></td></tr>
			       </table>
		                         <%
								 end if%>
								<textarea name="Content" ID="Content" style="display:none"><%=Server.HTMLEncode(Content)%></textarea>
                                <script type="text/javascript">CKEDITOR.replace('Content', {width:"98%",height:"320",toolbar:"Basic",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>
				             
	</td>
  </tr>
<%end if%>
<%if F_B_Arr(10)=1 Then%>	  
 <tr class="tdbg">
    <td height="25" align="center"><%=F_V_Arr(10)%>：</td>
    <td height="25"> <table width="100%">
		<tr>
		 <td><input name='PhotoUrl' type='text' id='PhotoUrl' value="<%=PhotoUrl%>" size='40'  class="textbox"/></td>
		 <td>
			<%if F_B_Arr(11)=1 Then%>	
              <iframe  id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=<%=ChannelID%>' frameborder="0" scrolling="No"  width='400' height='24'></iframe>
			<%end if%>   
		</td>
		</tr>
	</table>
	</td>
</tr>
<%end if%>
<%If F_B_Arr(18)=1 Then%>
   <tr class="tdbg">
        <td height="25" align="center"><span>阅读<%=KS.Setting(45)%>：</span></td>
         <td height="25"><input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6"><%=KS.Setting(46)%> <span class="msgtips">如果免费阅读请输入“<font color=red>0</font>”</span></td>
   </tr>
<%end if%>


 <tr class="tdbg">
   <td height="40"></td><td><button class="pn" id="submit1" type="submit"><strong>OK, 保 存</strong></button>&nbsp;<%if IsTemplate Then 
   Response.Write "[#Status]" 
   Elseif id<>0 Then
		     If Verific<>1 Then
			  if Verific=2 Then
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' checked>放入草稿</label>"
			  else
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' value='2'>放入草稿</label>"
			  end if
			 Else
			  response.write "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
	Else
		    response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2'>放入草稿</label>"
   End If%> </td>
 </tr>
</table>
<br/><%
End Function
%>