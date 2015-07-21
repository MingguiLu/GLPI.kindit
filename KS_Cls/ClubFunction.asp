<%
Sub Echo(sStr)
			If Immediate Then
				Response.Write    sStr
				Response.Flush()
			Else
				Templates    = Templates&sStr 
			End If 
End Sub 
		
Public Sub Scan(sTemplate)
			Dim iPosLast, iPosCur
			iPosLast    = 1
			While True 
				iPosCur    = InStr(iPosLast, sTemplate, "{@")
				If iPosCur>0 Then
					Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
					iPosLast    = Parse(sTemplate, iPosCur+2)
				Else 
					Echo    Mid(sTemplate, iPosLast)
					Exit Sub  
				End If 
		   Wend 
End Sub 

Sub GetClubPopLogin(ByRef FileContent)
 If Instr(FileContent,"{#GetClubPopLogin}")=0 Then Exit Sub
 Dim Str
 If KS.IsNul(KS.C("UserName")) And KS.IsNul(KS.C("PassWord")) Then
   Str="您好,欢迎进入" & KS.Setting(0) & "! [<a href=""javascript:void(0)"" onclick=""ShowLogin()"">登录</a>] | [<a href='" & KS.GetDomain & "?do=reg' target='_blank'>免费注册</a>]"
 Else
   Dim GetMailTips,MyMailTotal:MyMailTotal=GCls.Execute("Select Count(ID) From KS_Message Where Incept='" &KS.C("UserName") &"' And Flag=0 and IsSend=1 and delR=0")(0)
   IF MyMailTotal>0 Then 
	  GetMailTips="<span style=""color:red"">" & MyMailTotal & "</span><bgsound src=""" & KS.GetDomain & "User/images/mail.wav"" border=0>"  
   Else
	  GetMailTips=0
   End If
   Str="您好！<span style='color:red'>" & KS.C("UserName") & "</span>,欢迎来到会员中心!【<a href='" & KS.GetDomain & "user/'>会员中心</a>】【<a href='" & KS.GetDomain & "user/user_mytopic.asp?action=fav'>我收藏的帖子</a>】【<a href='" & KS.GetDomain & "user/user_Message.asp?action=inbox'>短消息"&GetMailTips&"</a>】【<a href='" & KS.GetDomain & "User/UserLogout.asp'>退出</a>】"
       Dim KSUser:Set KSUser=New UserCls
	KSUser.UserLoginChecked
    If KS.ChkClng(KS.U_S(KSUser.GroupID,8))>0 and KS.ChkClng(KS.U_S(KSUser.GroupID,9))>0 And datediff("n",KSUser.GetUserInfo("LastLoginTime"),now)>=KS.ChkClng(KS.U_S(KSUser.GroupID,8)) then '判断积分奖励时间
     str=str & "<script>popShowMessage('" & KS.Setting(3) & KS.Setting(66) &"','" & KS.U_S(KSUser.GroupID,8) & "分钟后重新登录，奖励积分 +" & KS.U_S(KSUser.GroupID,9) & "分！');</script>"
	 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(KS.U_S(KSUser.GroupID,9)),"系统",KS.ChkClng(KS.U_S(KSUser.GroupID,8)) & "分钟后,重新登录奖励获得",0,0)
	  Conn.Execute("Update KS_User Set LastLoginTime=" & SQLNowString & " Where UserName='" & KSUser.UserName & "'")
	  If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@lastlogintime").Text=now
	ElseIf Not KS.IsNUL(Session("PopTips")) Then
     str=str & "<script>popShowMessage('" & KS.Setting(3) & KS.Setting(66) &"','" & Session("PopTips") & "！');</script>"
	 Session("PopTips")=""
	End if
	Set KSUser=Nothing

 End If
   FileContent=Replace(FileContent,"{#GetClubPopLogin}",str)
End Sub

'取得所有置顶帖子
Sub LoadTopTopic()
  If Not IsObject(Application(KS.SiteSN &"TopXML")) Then
   MustReLoadTopTopic
  End If
End Sub
Sub MustReLoadTopTopic()
	  Dim ListTopicFields:ListTopicFields="ID,UserName,UserID,Subject,AddTime,Verific,LastReplayUser,LastReplayUserID,LastReplayTime,TotalReplay,BoardID,Hits,IsPic,IsTop,IsBest,PostType,AnnexExt,CategoryId" rem 主题列表用到的字段
	  Dim RS:Set RS=Conn.Execute("Select top 500 " & ListTopicFields & " From KS_GuestBook Where Verific<>0 And IsTop<>0 Order BY LastReplayTime Desc")
	  If Not RS.Eof Then
		Set Application(KS.SiteSN &"TopXML")=KS.RsToXml(RS,"row","")
	  End If
	 RS.Close:Set RS=Nothing
End Sub

Sub LoadMasterUserID(BoardID,ByVal Master)
  Dim Users,RS,str
  If Instr(Master,",")=0 Then 
   Users="'" & Master & "'"
  Else
   Users="'" & Replace(Master,",","','") &"'"
  End If
  Set RS=Conn.Execute("Select UserID,UserName From KS_User Where UserName in (" & Users & ")")
  If Not RS.Eof Then
    Do While Not RS.Eof
	  If Str="" Then
	    str=rs(0) & "|" & rs(1)
	  Else
	    str=str & "@" & rs(0) & "|" & rs(1)
	  End If
	RS.MoveNext
	Loop
  End If
  RS.Close : Set RS=Nothing
  Application(KS.SiteSN &"Master"&BoardID)=str
End Sub

'取得投票
Function GetVote(TopicID,Xml)
Dim rs,votestr,VNode,VoteType,VoteID,VoteN,TotalVote,VoteColorArr,CanVote,VoteNums,VoteUserList,TimeLimit,EndTime,IPnums
	VoteColorArr=Array("#E92725","#F27B21","#F2A61F","#5AAF4A","#42C4F5","#0099CC","#3365AE","#2A3591","#592D8E","#DB3191","#cccccc")
	votestr="<div id=""showvote""><table width=""550"" class=""votetable"" cellspacing=""0"" cellpadding=""0""><tr><td colspan=""2"">"
   set rs=conn.execute("select top 1 * from ks_vote where topicid=" & TopicID)
   if rs("VoteType")="Single" Then votestr=votestr & "<strong>单选投票</strong>" Else votestr=votestr &"<strong>多选投票</strong>"
   VoteType=RS("VoteType") :TimeLimit=Rs("TimeLimit") : EndTime=Rs("TimeEnd")
   VoteID=RS("ID") : VoteNums=RS("VoteNums") : VoteUserList=rs("VoteUserList")
   IPnums=RS("IpNums")
   RS.Close : Set RS=Nothing
   If IpNums=1 And KS.FoundInArr(VoteUserList,KS.C("UserName"),",")=true Then CanVote=false Else CanVote=True
   if TimeLimit="1" then votestr=votestr & ",结束时间:"& endtime
   votestr=votestr & ",共有" &VoteNums &"人参与投票, <a href=""javascript:void(0)"" onclick=""showVoteUser('"& KS.Setting(66) &"'," & VoteID& ")"">查看参与用户</a></td></tr>"
						  
						  If Not IsObject(XML) Then Set XML=LFCls.GetXMLFromFile("voteitem/vote_"&VoteID)
						  For Each VNode In Xml.DocumentElement.SelectNodes("voteitem")
							   TotalVote=TotalVote+KS.ChkClng(VNode.childNodes(1).text)
						  Next
						  VoteN=1
						  For Each VNode In Xml.DocumentElement.SelectNodes("voteitem")
						   votestr=votestr & "<tr><td height='30' colspan=""2"">"
						   If CanVote=True Then
							   If VoteType="Single" Then
							   votestr=votestr&"<label><input type='radio' name='VoteOption' value='"& VNode.getAttribute("id") &"' />"
							   Else
							   votestr=votestr&"<label><input type='checkbox' name='VoteOption' value='"& VNode.getAttribute("id") &"' />"
							   End If
						   End If
						   votestr=votestr&VoteN &"、" & VNode.childNodes(0).text & "</label>"
						   votestr=votestr &"</td></tr>"
						   
						   dim perVote,pstr,votebg
							if totalVote=0 Then TotalVote=0.00000001
							perVote=round(VNode.childNodes(1).text/totalVote,4)
							votebg=round(480*perVote)
							perVote=perVote*100
							if perVote<1 and perVote<>0 then
								pstr="&nbsp;0" & perVote & "%"
							else
								pstr="&nbsp;" & perVote & "%"
							end if
						   
						   votestr=votestr & "<tr><td><div class=""vbg""><div style=""width:" & votebg & "px;background:" & VoteColorArr(voten-1) &""">&nbsp;</div></div></td><td align=""left"">" &pstr&"<em style=""color:" & VoteColorArr(voten-1) &""">(" & VNode.childNodes(1).text &")</em></td></tr>"
						   VoteN=VoteN+1
						  Next
						  votestr=votestr &"<tr><td style=""height:40px"" colspan=""2"">"
						  If CanVote Then
						  votestr=votestr&"<input type=""button"" onclick=""doVote('" & KS.Setting(66) & "'," & VoteID & ",'" & VoteType &"')"" id=""votebtn"" value=""投票"" />"
						  Else
						  votestr=votestr&"<input type=""button"" disabled id=""votebtn"" value=""投票"" />"
						  End If
						  votestr=votestr& "</td></tr>"
						VoteStr=VoteStr & "</table></div>"
		GetVote=VoteStr
End Function
'取得点评
Function GetComments(CommentXML,BoardID,replayid,MaxPerPage,IsMaster)
     Dim Str,j,N,P,PageNum,TotalPut
	 P=KS.ChkClng(KS.S("P")) : If P<=0 Then P=1
	 N=0
     If IsObject(CommentXML) Then
		Dim UserFace,CN,CMT,CommentNodes:Set CommentNodes=CommentXML.DocumentElement.SelectNodes("row[@pid=" & replayid & "]")
		TotalPut=CommentNodes.length
		If TotalPut>0 Then
			    if (TotalPut mod MaxPerPage)=0 then
				    PageNum = TotalPut \ MaxPerPage
				else
					PageNum = TotalPut \ MaxPerPage + 1
				end if
				If P>PageNum Then P=PageNum

				Str= "<h3>点评 <span>共 <span class='red'>" &TotalPut& "</span> 条</span></h3>"
				For J=0 To TotalPut
				    Set CN=CommentNodes.Item((p-1)*MaxPerPage+n)
					If CN Is Nothing Then Exit For
			  ' For Each CN In CommentNodes
					CMT=replace(cn.selectsinglenode("@comment").text,chr(10),"<br/>")
					Str= Str & "<div class=""pstl"
					If TotalPut>1 Then Str=Str &" line"
					Str=Str & """>"
					If CN.SelectSingleNode("@userid").text="0" And Instr(CMT,"：")<>0 Then
							Dim K,KK,GD,Star,CommentArr:CommentArr=Split(CMT,"：")
							For K=0 To Ubound(CommentArr)-1
								If K=0 Then
									  GD=CommentArr(k)
								ElseIf Instr(CommentArr(k),"</i> ")<>0 Then
									  GD=split(CommentArr(k),"</i> ")(1)
								End If
								Star=KS.CutFixContent(CMT,GD&"：<i>","</i>",0)
								Str= Str &  GD & "：<span class='red'>" & formatnumber(star,1,-1,-1) & "</span> "
								For KK=0 To 4
								  if cint(kk+1)<=cint(star) Then
									 Str= Str & "<span class='currstar' title='" & star &"'>★</span>"
								  Else
									 Str= Str & "<span class='star'>★</span>"
								  End If
								Next
									Str= Str & "&nbsp;&nbsp;&nbsp;&nbsp;"
							Next
							Str= Str & "</div>"
					Else
							 UserFace=CN.SelectSingleNode("@userface").text
							 If Not KS.IsNUL(UserFace) Then
								If Left(UserFace,1)<>"/" And Left(lcase(UserFace),4)<>"http" Then UserFace=KS.GetDomain & UserFace
							 Else
								UserFace=KS.GetDomain & "images/face/boy.jpg"
							 End If 
							 Str= Str & "<div class=""psta""><a href=""" & KS.GetSpaceUrl(cn.selectsinglenode("@userid").text) & """ target=""_blank""><img onerror='this.src=""" & KS.Setting(3) & "images/face/boy.jpg""' src=""" & UserFace & """ /></a></div>"
							 Str= Str & "<div class=""psti"">"
							 Str= Str & "<a href=""" & KS.GetSpaceUrl(cn.selectsinglenode("@userid").text) & """ target=""_blank"">" & cn.selectsinglenode("@username").text &"</a>&nbsp;" & CMT & "&nbsp;"
										
							 dim ps:ps=cn.selectsinglenode("@prestige").text
							 If KS.ChkClng(ps)<>0 Then 
								   if ps>0 Then
										Str= Str & "威望<span class=""ww"">+" & ps &"</span>&nbsp;"
								   else
										Str= Str & "威望<span class=""ww"">" & ps &"</span>&nbsp;"
								   end if
							 end if
							 Str= Str & "<span class=""xg1"">发表于 " & KS.GetTimeFormat1(cn.selectsinglenode("@adddate").text,true) & "&nbsp;</span>"
							 If IsMaster Then str=str &" <a href='javascript:void(0)' onclick='delCmt(""" & KS.Setting(66) & """," & CN.SelectSingleNode("@id").text & "," & ReplayID&"," & BoardID&"," & p & ")'>删除</a>"
							 Str= Str & "</div>"
							Str= Str & "</div>"
					End If
					N=N+1
					If N>=MaxPerPage Then Exit For
			   Next
			   
		  Str=Str &"<div class=""cmtpage"">"
		  If PageNum>1 Then
			  If P=1 Then 
			   Str=Str &"<a href='javascript:void(0)' onclick='ShowCmtPage(""" & KS.Setting(66) & """,2," & replayid&"," & BoardID&")'>下一页 >> </a>"   
			  Else
				  If P>1 Then
				  Str=Str &"<a href=javascript:void(0) onclick=ShowCmtPage('" & KS.Setting(66) & "',"&p-1& "," & replayid&"," & BoardID&")><< 上一页</a>"   
				  End If
				For K=1 To PageNum
				  If K=P Then
				  Str=Str &"<a href='javascript:void(0)' class='curr'>" & k & "</a>"   
				  Else
				  Str=Str &"<a href=javascript:void(0) onclick=ShowCmtPage('" & KS.Setting(66) & "',"&k& "," & replayid&"," & BoardID&")>" & k & "</a>"   
				  End If
				Next
				If P>1 And P<>PageNum Then
				  Str=Str &"<a href=javascript:void(0) onclick=ShowCmtPage('" & KS.Setting(66) & "',"&p+1& "," & replayid&"," & BoardID&")>下一页 >></a>"   
				End If
			  End If
		  End If
		 Str=Str &"</div>"
			   
		  End If
	 End If
	 GetComments=str
End Function
%>