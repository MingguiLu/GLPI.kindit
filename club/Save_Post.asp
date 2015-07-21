<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Config.Club.asp"-->

<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Guest_SaveData
KSCls.Kesion()
Set KSCls = Nothing

Class Guest_SaveData
        Private KS,KSUser,Node,LoginTF,FieldRndID,TopicID,UserID
        Private UserName, Email, Subject, Oicq, Verifycode, IP, Pic, TxtHead, HomePage, Content, ErrorMsg, a,BoardID,Purview,ShowIP,ShowSign,ShowScore,CategoryId,PopTips,posttype,VoteItemArr,VoteNum,VoteNumArr,voteitem,ValidDays,TimeBegin,TimeEnd,voteid,i
		Private O_LastPost,N_LastPost,O_LastPost_A,BSetting,Master
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
	   %>
	   <!--#include file="../KS_Cls/ClubFunction.asp"-->
	   <%
	   Public Sub Kesion()
		Dim TmpIsSelfRefer,I,SplitStrArr
		TmpIsSelfRefer = IsSelfRefer()
			
		If TmpIsSelfRefer <> TRUE Then '外部提交的数据
			Call KS.Alert("数据提交错误!", "")
			Exit Sub
		End If
		If Request.servervariables("REQUEST_METHOD") <> "POST" Then
			Response.Write "<script>alert('请不要非法提交！');</script>"
			Response.end
		End If
		If KS.IsNul(Request.ServerVariables("HTTP_REFERER")) Then
			Response.Write "<script>alert('请不要非法提交！');</script>"
			Response.end
		End If
		if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"post.asp")<=0 then
			Response.Write "<script>alert('非法提交！');</script>"
			Response.end
		end if

		
		LoginTF=KSUser.UserLoginChecked
		If KS.ChkClng(KS.C("UserID"))<>0 Then
			UserID = KS.ChkClng(KS.C("UserID"))
		Else
			UserID = KS.ChkClng(KSUser.GetUserInfo("userid"))
		End If
		
	   If KS.Setting(57)="1" and LoginTF=false Then
	     Call KS.Alert("没有发表权限!", "")
		 Exit Sub
	   End If
		
		FieldRndID=Session("Rnd")
		If KS.IsNul(FieldRndID) Then
	     Call KS.Alert("会话超时，请重新打开发帖窗口再提交!", "")
		 Exit Sub
		End If
		if mid(KS.Setting(161),3,1)="1" Then
			If KS.IsNul(Session("Qid")) Then
			 Call KS.Alert("会话超时，请重新打开发帖窗口再提交!", "")
			 Exit Sub
			Else
			 If Lcase(Request.Form("Answer" & FieldRndID))<>Lcase(Split(KS.Setting(163),vbcrlf)(KS.ChkClng(Session("Qid")))) Then
				 KS.Die "<script>alert('对不起，您的回答不正确!');</script>"
				 Exit Sub
			 End If
			End If
		End If
		
		
		Dim LastLoginIP:LastLoginIP = KS.GetIP
			UserName = KS.S("Name")
			Email = KS.S("Email")
			HomePage = KS.S("HomePage")
			Oicq = KS.ChkClng(KS.S("Oicq"))
			Verifycode = KS.S("Code"&FieldRndID)
			IP = LastLoginIP
			Pic = KS.S("Pic")
			TxtHead = KS.S("txthead")
			Subject = KS.S("Subject"&FieldRndID)
			posttype=KS.ChkClng(KS.S("posttype"))
			If posttype=1 Then  '投票
			 voteitem=KS.S("voteitem")
			 If KS.IsNul(voteitem) Then
				 KS.Die "<script>alert('对不起，投票帖必须输入投票选项!');</script>"
				 Exit Sub
			 End If
			 VoteItemArr=Split(voteitem,",")
			 If Ubound(VoteItemArr)<1 Then
				 KS.Die "<script>alert('对不起，投票选项不能少于两项!');</script>"
				 Exit Sub
			 End If
			 ValidDays=KS.ChkClng(Request.Form("ValidDays"))
			 If KS.S("timelimit")="1" And ValidDays<=0 Then
				 KS.Die "<script>alert('对不起，有效天数必须大于0!');</script>"
				 Exit Sub
			 End If
			 TimeBegin=now
			 TimeEnd=dateadd("d",ValidDays,now)
			End If
			
			
			Content = Request.Form("Content")
			Content=replace(Content,chr(10),"[br]")
			'非管理员及版主过滤标题html
			If KSUser.GetUserInfo("ClubSpecialPower")="0"  Then
			 Subject=KS.LoseHtml(Subject)
			End If
			Content=Server.HTMLEncode(Content)
			BoardID=KS.ChkClng(KS.S("BoardID"))
			Purview=KS.ChkClng(Request.Form("purview"))
			showip=KS.ChkClng(Request.Form("showip"))
			showsign=KS.ChkClng(Request.Form("showsign"))
			showscore=KS.ChkClng(Request.Form("showscore"))
			CategoryId=KS.ChkClng(Request.Form("CategoryId"))
			Content=KS.FilterIllegalChar(Content)
			'防发帖机
            dim kk,sarr
            sarr=split(WordFilter,"|")
            for kk=0 to ubound(sarr)
               if instr(content,sarr(kk))<>0 or instr(Subject,sarr(kk))<>0 then 
                  ks.die "<script>alert('含有非常关键词:" & sarr(kk) &",请不要非法提交恶意信息！');</script>"
               end if
            next
		a = CheckEnter()
		If Content="" Then
		 a=false
		 ErrorMsg="发表内容不能为空！"
		End If
		If a = True Then 
		    If BoardID<>0 Then
			 KS.LoadClubBoard()
			 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 O_LastPost=Node.SelectSingleNode("@lastpost").text
			 BSetting=Node.SelectSingleNode("@settings").text
			 Master=Node.SelectSingleNode("@master").text
			End If
			BSetting=BSetting&"$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
			BSetting=Split(BSetting,"$")
			
			If KS.ChkCLng(BSetting(40))<>0 Then
			  If len(replace(replace(KS.LoseHtml(Content),"	",""),vbcrlf,""))<KS.ChkCLng(BSetting(40)) Then
				Call KS.Alert("内容字数不能少于" &KS.ChkCLng(BSetting(40)) & "个字节!" , "")
				Response.End
			  End If
			End If
			
		     
			If KS.S("Action")="edit" Then
			 EditSave()
			Else 
			 SaveData()
			End If
			
			If KS.Setting(52)="1" Then   '帖子需要审核
			    Response.Write("<script>alert('发布成功,您发表的主题审核后才会显示！');top.location.href='" & KS.GetClubListUrl(boardid) & "';</script>")
			Else
				Session("PopTips")=PopTips
				Response.Write("<script>top.location.href='" & KS.GetClubShowUrl(TopicID)& "';</script>")
			End If
		Else
	        Call KS.Alert(ErrorMsg, "")
			Response.End
		End If
	
	End Sub
	
	Function CheckEnter()
	        If KS.C("UserName")="" then
			  UserName="游客：" & UserName
			Else
			  UserName=KS.C("UserName")
			end if
			IF Trim(Verifycode)<>Trim(Session("Verifycode")) And KS.ChkClng(KS.Setting(53))=1 then 
		   	 CheckEnter=False
			 ErrorMsg="验证码有误，请重新输入！"
			Else
			    If Subject="" Then
				   CheckEnter=False
				   ErrorMsg="请填写主题！"
				End If
				
				If KS.Setting(59)="1" Then 
					If UserName="" Then
						CheckEnter=False
						ErrorMsg="你好像忘了填“昵称”！"
					Else
						If Email="" or InStr(2,Email,"@")=0 Then
							CheckEnter=False
							ErrorMsg="你的Email有问题请重新填写！"
						Else
								If TxtHead="" Then
									CheckEnter=False
									ErrorMsg="你的表情没选！"
								Else
									If replace(Content,"&nbsp;","")="" Then
										CheckEnter=False
										ErrorMsg="留言不能为空！"
									Else
										CheckEnter=TRUE
									End If
								End If
						End If	   
					End If
				Else
				  CheckEnter=TRUE
				End If
			End If
		End Function
		
		'新增保存
		Sub SaveData()
			if datediff("n",KSUser.GetUserInfo("RegDate"),now)<KS.ChkClng(bsetting(9)) Then
				KS.Die "<script>alert('对不起,本版面限制" & bsetting(9) & "分钟内注册的会员不能发帖!');</script>"
			End if
			If (Not KS.IsNul(BSetting(2)) Or KS.ChkCLng(BSetting(3))<0) And LoginTF=false Then
				KS.Die "<script>alert('对不起,请先登录!');parent.ShowLogin()</script>"
			End If
			If KS.ChkCLng(BSetting(3))<0 And KS.ChkCLng(KSUser.GetUserInfo("Score"))<-KS.ChkCLng(BSetting(3)) Then
				KS.Die "<script>alert('对不起,在此版面发帖的需要清耗" & -KS.ChkCLng(BSetting(3)) & "分的积分,您当然积分余额为" & KSUser.GetUserInfo("Score") & "分不足以支付!');</script>"
			End If
			
			If KS.ChkClng(BSetting(41))<>0 Then
             If IsDate(Session(KS.SiteSN & "posttime"))  Then
				If DateDiff("s",Session(KS.SiteSN & "posttime"),Now())<KS.ChkClng(BSetting(41)) Then
					KS.Die "<script>alert('对不起,此版面设定发帖间隔时间不能少于" & BSetting(41)& "秒!');</script>"
				End If
			 End If
			 Session(KS.SiteSN & "posttime")=Now()
			End If
						
			 Dim GroupPurview:GroupPurview= True : If Not KS.IsNul(BSetting(1)) and KS.FoundInArr(Replace(BSetting(1)," ",""),KSUser.GroupID,",")=false Then GroupPurview=false
			Dim UserPurview:UserPurview=True : If Not KS.IsNul(BSetting(10)) and KS.FoundInArr(BSetting(10),KSUser.UserName,",")=false Then UserPurview=false
			Dim ScorePurview:ScorePurview=KS.ChkClng(BSetting(11))
			Dim MoneyPurview:MoneyPurview=KS.ChkClng(BSetting(12))
			Dim LimitPostNum:LimitPostNum=KS.ChkClng(BSetting(13))
			
			If KS.Setting(59)<>"1" Then
				If KSUser.GetUserInfo("LockOnClub")="1" Then
					KS.Die "<script>alert('对不起,您的账号在本论坛被锁定,无权发帖!');</script>"
				ElseIf (GroupPurview=false and Not KS.IsNul(BSetting(10))) or (UserPurview=false) Then
					KS.Die "<script>alert('对不起,您没有在此版面发帖的权限!');</script>"
				ElseIf (ScorePurView>0 and KS.ChkClng(KSUser.GetUserInfo("Score"))<ScorePurView) Or (MoneyPurview>0 and KS.ChkClng(KSUser.GetUserInfo("Money"))<MoneyPurview) Then
					KS.Die "<script>alert('对不起,您积分或资金不足!');</script>"
				ElseIf LimitPostNum<>0 Then
						 Dim PostNum:PostNum=Conn.Execute("Select count(1) From KS_GuestBook Where BoardId=" & BoardID & " and UserName='" & KSUser.UserName &"' And DateDiff(" & DataPart_D & ",AddTime," & SqlNowString & ")<1")(0)
						 If PostNum>=LimitPostNum Then
							  KS.Die "<script>alert('对不起,本版面每天限制发表" & limitpostnum & "个主题!');</script>"
						 End If
				End If
			End If
			
			If KS.IsNul(Subject) Then Subject=Left(KS.LoseHtml(Content),100)
			
			'取帖子存放数据表
			Dim Nodes,Doc,TableName
			set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
			Set Nodes=Doc.DocumentElement.SelectSingleNode("item[@isdefault='1']")
			TableName=nodes.selectsinglenode("tablename").text
			Set Doc=Nothing


		    Dim SqlStr:SqlStr = "SELECT top 1 * From KS_GuestBook WHERE 1=0" 
			Dim RSObj:Set RSObj=Server.CreateObject("Adodb.RecordSet")
			RSObj.Open SqlStr,Conn,1,3
			RSObj.AddNew 
			RSObj("PostTable")=TableName
			RSObj("UserName") = KS.HTMLEncode(UserName)
			RSObj("UserID") = UserID
			RSObj("Email") = KS.HTMLEncode(Email)
			RSObj("HomePage") = KS.HTMLEncode(HomePage)
			if KS.Setting(59)="0" then
			 RSObj("Face") =Pic
			 If Not KS.IsNul(Pic) Then
			  If lcase(Right(pic,"3"))="gif" Then
			  RSObj("isPic")=1   'gif
			  Else
			  RSObj("isPic")=2    'jpg
			  End If
			 Else
			  RSObj("IsPic")=0
			 End If
			else
			If KSUser.GetUserInfo("Sex")="男" Then RSObj("Face") ="boy.jpg" Else  RSObj("Face") ="girl.jpg"
			 RSObj("IsPic")=0
			end if
			RSObj("TxtHead") = TxtHead&".gif"
			RSObj("Subject") = Subject
			'RSObj("Content") = Content
			RSObj("Oicq") =Oicq       
			RSObj("GuestIP") = IP  
			If KS.Setting(52)="1" Then  
			RSObj("Verific")=0
			Else
			RSObj("Verific")=1
			End If
			RSObj("AddTime") = Now()
			RSObj("Hits")=0
			RSObj("IsTop")=0
			RSObj("IsBest")=0
			RSObj("IsSlide")=0
			RSObj("DelTF")=0
			RSObj("BoardID")=BoardID
			RSObj("Purview")=Purview
			RSObj("ShowIP")=ShowIP
			RSObj("ShowSign")=ShowSign
			RSObj("ShowScore")=ShowScore
			RSObj("CategoryId")=CategoryId
			RSObj("LastReplayTime")=Now
			RSObj("TotalReplay")=0
			RSObj("LastReplayUser")=KS.HTMLEncode(UserName)
			RSObj("LastReplayUserID")=UserID
			RSObj("AnnexExt")=KS.S("AnnexExt")
			RSObj("posttype")=posttype
			RSObj.Update
			RSObj.MoveLast
			TopicID=RSObj("ID")
			N_LastPost=RSObj("ID")&"$"& now & "$" & Replace(left(subject,200),"$","") & "$" & UserName & "$" &UserID&"$$"
			RSObj.Close
			
			'写入到回复表
			SqlStr = "SELECT top 1 * From " & TableName &" WHERE ID IS NULL" 
			RSObj.Open SqlStr,Conn,1,3
			RSObj.AddNew 
			RSObj("UserName") = UserName
			RSObj("UserID") = UserID
			RSObj("UserIP") = KS.GetIP
			RSObj("TopicID") = TopicID
			RSObj("Content") =Content
			RSObj("TxtHead")=TxtHead
			RSObj("ShowIp")=ShowIP
			RSObj("ShowSign")=ShowSign
			RSObj("ReplayTime") = Now
			RSObj("ParentId")=0
			If KS.Setting(52)="1" Then  
			RSObj("Verific")=0
			Else
			RSObj("Verific")=1
			End If
			RSObj("DelTF")=0
			RSObj.Update
			RSObj.Close
			
			
			If posttype=1 Then   '投票
					rsobj.open "select top 1 * from KS_Vote",conn,1,3
					rsobj.addnew
						 rsobj("TopicID")=TopicID
						 rsobj("Title")=Subject
						 rsobj("timelimit")=KS.ChkClng(KS.G("TimeLimit"))
						 rsobj("TimeBegin")=TimeBegin
						 rsobj("TimeEnd")=TimeEnd
						 rsobj("nmtp")=KS.ChkClng(Request("nmtp"))
						 rsobj("groupids")=""
						 rsobj("ipnum")=1
						 rsobj("ipnums")=1
						 rsobj("templateid")="{@TemplateDir}/投票页.html"
						 rsobj("status")=1
						 rsobj("AddDate")=Now
						 rsobj("VoteType")=KS.S("VoteType")
						 rsobj("UserName")=UserName
						 rsobj("NewestTF")=0
						 rsobj("VoteNums")=0
					 rsobj.update
					 rsobj.movelast
					 voteid=rsobj("id")
					 rsobj.close
					
					Dim XMLStr:XMLStr="<?xml version=""1.0"" encoding=""gb2312"" ?>" &vbcrlf
					XMLStr=XMLStr&" <vote>" &vbcrlf
					for i=0 to ubound(VoteItemArr)
					  if trim(VoteItemArr(i))<>"" Then
					    XMLStr=XMLStr & "  <voteitem id=""" & i+1 &""">"&vbcrlf
						XMLStr=XMLStr & "    <name>" & VoteItemArr(i) &"</name>" &vbcrlf
						XMLStr=XMLStr & "    <num>0</num>" &vbcrlf
					    XMLStr=XMLStr & "  </voteitem>"&vbcrlf
					  End If
					Next
					XMLStr=XMLStr &" </vote>" &vbcrlf
					Call KS.WriteTOFile(KS.Setting(3) & "config/voteitem/vote_" & voteid & ".xml",xmlstr)
			        Application(KS.SiteSN&"_Configvoteitem/vote_"&VoteID)=empty
			End If
			
			Set RSObj = Nothing
			
			Session("UploadClassID")=BoardID
            If Not KS.IsNul(Session("UploadFileIDs")) Then 
				 Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & TopicID &",classID=" & BoardID & " Where ID In (" & KS.FilterIds(Session("UploadFileIDs")) & ")")
			End If
			If LoginTF=true Then
			  Conn.Execute("Update KS_User Set PostNum=PostNum+1 Where UserName='" & KSUser.UserName & "'")
			End If
			
			'关联上传文件
			Call KS.FileAssociation(9994,TopicID,Content,0)
			
			If KS.ChkClng(BSetting(3))>0 and LoginTF=true Then
			    PopTips="<strong>积分" & KSUser.GetUserInfo("Score") &"+</strong>" & KS.ChkClng(BSetting(3))
				Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(3)),"系统","在论坛发表主题[" & Subject & "]所得!",0,0)
			End If
			If KS.ChkClng(BSetting(3))<0 and LoginTF=true Then
			    PopTips="<strong>积分" & KSUser.GetUserInfo("Score") &"-</strong>" & -KS.ChkClng(BSetting(3))
				Call KS.ScoreInOrOut(KSUser.UserName,2,-KS.ChkClng(BSetting(3)),"系统","在论坛发表主题[" & Subject & "]消费!",0,0)
			End If
			If LoginTF=true Then
			  If KS.ChkClng(BSetting(30))<>0 Then
			  if PopTips="" then
			   PopTips="<strong>威望" & KSUser.GetUserInfo("Prestige") &"+</strong>" & -KS.ChkClng(BSetting(30))
			  Else
			   PopTips=PopTips & ",<strong>威望" & KSUser.GetUserInfo("Prestige") &"+</strong>" & KS.ChkClng(BSetting(30))
			  end if
			  If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@prestige").Text=KS.ChkClng(KSUser.GetUserInfo("Prestige"))+KS.ChkClng(BSetting(30))
			  Conn.Execute("Update KS_User Set Prestige=Prestige+" & KS.ChkClng(BSetting(30)) & " Where UserName='" & KSUser.UserName &"'")
			  End If
			  Call KSUser.AddLog(KSUser.UserName,"在论坛发表了主题[<a href='{$GetSiteUrl}club/display.asp?id=" & TopicID & "' target='_blank'>" & subject &"</a>]",100)
			End If
			
			'更新今日发帖数等
			If BoardID<>0 Then
			    If KS.Setting(52)="1" Then   '帖子需要审核
				Conn.Execute("Update KS_GuestBoard set postnum=postnum+1,topicnum=topicnum+1 where id=" & BoardID)
				Else
				Conn.Execute("Update KS_GuestBoard set lastpost='" & N_LastPost & "',postnum=postnum+1,topicnum=topicnum+1 where id=" & BoardID)
				End If
				If KS.IsNul(O_LastPost) Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				Else
				 O_LastPost_A=Split(O_LastPost,"$")
				 Dim LastPostTime:LastPostTime=O_LastPost_A(1)
				 If Not IsDate(LastPostTime) Then LastPostTime=now
				 If datediff("d",LastPostTime,Now())=0 Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=todaynum+1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text)+1
				 Else
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				 End If
				End If
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text)+1
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@topicnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@topicnum").text)+1
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text=N_LastPost
		   End  If
			
			set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
					If DateDiff("d",xmldate,now)=0 Then
					  doc.documentElement.attributes.getNamedItem("todaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text+1
					  If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) then
					   doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  end if
					  
					Else
					  doc.documentElement.attributes.getNamedItem("date").text=now
					  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  doc.documentElement.attributes.getNamedItem("todaynum").text=0
					End If
					  doc.documentElement.attributes.getNamedItem("topicnum").text=doc.documentElement.attributes.getNamedItem("topicnum").text+1
					  doc.documentElement.attributes.getNamedItem("postnum").text=doc.documentElement.attributes.getNamedItem("postnum").text+1
					  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		End sub
		
		'修改保存数据
		Sub EditSave
		 Dim TopicID:TopicID=KS.ChkClng(KS.S("TopicID"))
		 Dim ReplyID:ReplyID=KS.ChkClng(KS.S("replyId"))
		 Dim IsTopic:IsTopic=KS.ChkClng(KS.S("IsTopic"))
		 Dim IsTop,Page:Page=KS.ChkClng(KS.S("Page"))
		 If Page=0 Then Page=1
		 Dim PostTable,PostUserName
		 if TopicID=0 Or ReplyID=0 Then
			 KS.Die "<script>alert('参数出错!');</script>"
		 End If
		 
		 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 PostTable,IsTop From KS_GuestBook Where ID=" & TopicID,conn,1,1
		 If RS.Eof And RS.Bof Then
		    RS.Close : Set RS=Nothing
		    KS.Die "<script>alert('参数出错!');</script>"
		 End If
		    PostTable=RS(0)
			IsTop=RS(1)
		 RS.Close
		 RS.Open "Select top 1 * From " & PostTable  & " Where ID=" & ReplyID,conn,1,3
	     If RS.Eof And RS.Bof Then
		    RS.Close : Set RS=Nothing
		    KS.Die "<script>alert('参数出错!');</script>"
		  End If
		  PostUserName=RS("UserName")
		  
		  '检查编辑权限
		  If CheckIsMater=false Then
			If KSUser.UserName<>PostUserName Or KS.ChkClng(BSetting(29))=0 Then
			  RS.Close :Set RS=Nothing
			  KS.Die "<script>alert('对不起,您没有修改帖子权限!');</script>"
			End If
		  End If
		  RS("Content")=Content
		  RS.Update
		  RS.Close:Set RS=Nothing
		  If IsTopic=1 Then
		     If PostType=1 Then
			        VoteNum=KS.S("VoteNum") &",0,0,0,0,0,0,0,0,0,0,0,0"
					VoteNumArr=Split(VoteNum,",")
			        Dim RSObj:Set RSObj=Server.CreateObject("adodb.recordset")
			        rsobj.open "select top 1 * from KS_Vote Where TopicID=" &TopicID ,conn,1,3
					If Not rsobj.eof Then
						 rsobj("Title")=Subject
						 rsobj("timelimit")=KS.ChkClng(KS.G("TimeLimit"))
						 rsobj("TimeBegin")=TimeBegin
						 rsobj("TimeEnd")=TimeEnd
						 rsobj("nmtp")=KS.ChkClng(Request("nmtp"))
						 rsobj("VoteType")=KS.S("VoteType")
					 rsobj.update
					 rsobj.movelast
					 voteid=rsobj("id")
					
					Dim XMLStr:XMLStr="<?xml version=""1.0"" encoding=""gb2312"" ?>" &vbcrlf
					XMLStr=XMLStr&" <vote>" &vbcrlf
					for i=0 to ubound(VoteItemArr)
					  if trim(VoteItemArr(i))<>"" Then
					    XMLStr=XMLStr & "  <voteitem id=""" & i+1 &""">"&vbcrlf
						XMLStr=XMLStr & "    <name>" & VoteItemArr(i) &"</name>" &vbcrlf
						XMLStr=XMLStr & "    <num>" & VoteNumArr(i) & "</num>" &vbcrlf
					    XMLStr=XMLStr & "  </voteitem>"&vbcrlf
					  End If
					Next
					XMLStr=XMLStr &" </vote>" &vbcrlf
					Call KS.WriteTOFile(KS.Setting(3) & "config/voteitem/vote_" & voteid & ".xml",xmlstr)
			        Application(KS.SiteSN&"_Configvoteitem/vote_"&VoteID)=empty
				End If
				rsobj.close : Set RSObj=Nothing
			 End If
		  
		    Conn.Execute("Update KS_GuestBook Set Subject='" & Subject & "',categoryid=" & KS.ChkClng(KS.S("CategoryID")) &" Where ID=" & TopicID)
			Call KS.FileAssociation(1036,ReplyID,Content,1)
		  Else
		    Call KS.FileAssociation(1035,ReplyID,Content,0)
		  End If
          If IsTop<>0 Then MustReLoadTopTopic
       
		  
          KS.Die "<script>top.location.href='" & KS.GetClubShowUrlPage(TopicId,Page) & "';</script>"
		End Sub
		
		
		'检查版主或管理员
       function CheckIsMater()
	    If Cbool(LoginTF)=false Then
		  CheckIsMater=false : Exit Function
		Elseif KSUser.GetUserInfo("ClubSpecialPower")=1 Or KSUser.GetUserInfo("ClubSpecialPower")=2 Or KSUser.GroupID=1 Then
		  CheckIsMater=true : Exit function
		else
		  CheckIsMater=KS.FoundInArr(master, KSUser.UserName, ",")
		end if
       End function

		
		' ============================================
		' 检测上页是否从本站提交
		' 返回:True,False
		' ============================================
		Function IsSelfRefer()
			Dim sHttp_Referer, sServer_Name
			sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER"))
			sServer_Name = CStr(Request.ServerVariables("SERVER_NAME"))
			If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then
				IsSelfRefer = True
			Else
				IsSelfRefer = False
			End If
		End Function
End Class
%> 
