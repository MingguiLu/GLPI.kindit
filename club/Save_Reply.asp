<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/Kesion.ClubCls.asp"-->
<!--#include file="Config.Club.asp"-->
<!--#include file="../Plus/Session.asp"-->
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
        Private KS,KSUser,Node,BSetting,PostTable,Rid,TopicNode,PopTips,UserID
        Private UserName,Subject, Verifycode,TxtHead, Content, ErrorMsg,TopicID,BoardID,LoginTF,ShowIP,ShowSign
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
		    If TmpIsSelfRefer <> TRUE Then 	KS.Die escape("error|数据提交错误！")
		    LoginTF=KSUser.UserLoginChecked
			If KS.Setting(54)<>"3" And LoginTF=false Then
			 KS.Die escape("error|对不起，你没有发表的权限！")
			ElseIf KSUser.GetUserInfo("LockOnClub")="1" Then
			 KS.Die escape("error|对不起，你的账号被锁定,无法回帖!")
			ElseIf KS.Setting(54)=1 And KSUser.GroupID<>1 Then
			 KS.Die escape("error|对不起，本站只允许管理人员回复!")
			ElseIf KS.Setting(54)=2 And LoginTF=False Then
			KS.Die escape("error|对不起，本站至少要求是会员才可以发表回复！")
			End If
			If KS.Setting(54)<>"3" And LoginTF=false Then KS.Die escape("error|没有发表权限!")
			BoardID=KS.ChkClng(Request("BoardID"))
			If BoardID<>0 Then
			 KS.LoadClubBoard()
			 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
			End If
			BSetting=BSetting & "$$$0$0$0$0$0$0$1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
			BSetting=Split(BSetting,"$")
			If Not KS.IsNul(BSetting(2)) And KS.FoundInArr(BSetting(2),KSUser.GroupID,",")=false Then KS.Die escape("error|你所在的用户组,没有发表权限!")
			
			If LoginTF= True Then UserName=KSUser.UserName Else UserName="游客"
			TopicID = KS.ChkClng(KS.S("TopicID"))
			If KS.ChkClng(KS.C("UserID"))<>0 Then
			    UserID = KS.ChkClng(KS.C("UserID"))
			Else
				UserID = KS.ChkClng(KSUser.GetUserInfo("userid"))
			End If
			Content = UnEscape(Request.Form("Content"))
			If KS.IsNul(Content) Then KS.Die escape("error|回复字数必须录入!")
			If len(replace(replace(KS.LoseHtml(Content),"	",""),vbcrlf,""))<KS.ChkCLng(BSetting(40)) Then KS.Die escape("error|回复字数不能少于" & KS.ChkCLng(BSetting(40)) & "个字符!")

			Content=replace(Content,chr(10),"[br]")
            Content=Server.HTMLEncode(Content)
			Content=KS.CheckScript(content)
			ShowIP=KS.ChkClng(Request("showip"))
			ShowSign=KS.ChkClng(Request("showsign"))
			TxtHead = KS.S("TxtHead")
			Content=KS.FilterIllegalChar(Content)
			PostTable=KS.S("PostTable")
			If PostTable="" Then KS.Die escape("error|非法参数！")
			If lcase(left(PostTable,8))<>"ks_guest" Then KS.Die escape("error|非法参数！")
		    If TopicID=0 Then KS.Die escape("error|非法参数！")
	        If Content="" Then KS.Die escape("error|你没有输入回复内容!")
			If KS.ChkClng(BSetting(14))=0 Then   '判断是否回复自己的帖子
			  If Conn.Execute("Select top 1 UserName From KS_GuestBook Where ID=" & TopicID)(0)=UserName And UserName<>"游客" Then
			  KS.Die escape("error|本版面设置会员不能回复自己的主题帖!")
			  End If
			End If
			
			'防发帖机
            dim kk,sarr
            sarr=split(WordFilter,"|")
            for kk=0 to ubound(sarr)
               if instr(Content,sarr(kk))<>0 then 
                  ks.die  escape("error|含有非常关键词:" & sarr(kk) &",请不要非法提交恶意信息！")
               end if
            next
			
			If KS.ChkClng(BSetting(41))<>0 Then
             If IsDate(Session(KS.SiteSN & "posttime"))  Then
				If DateDiff("s",Session(KS.SiteSN & "posttime"),Now())<KS.ChkClng(BSetting(41)) Then
				   KS.Die escape("error|请休息下稍候再灌,此版面设定发帖间隔时间不能少于" & BSetting(41)& "秒!")
				End If
			 End If
			 Session(KS.SiteSN & "posttime")=Now()
			End If

			SaveData
			If KS.ChkClng(KS.S("IsTop"))<>0 Then MustReLoadTopTopic
			If KS.ChkClng(KS.S("toend"))=1 Then
			 Dim MaxPerPage:MaxPerPage=KS.ChkClng(BSetting(21)) : If MaxPerPage=0 Then MaxPerPage=10
			 Dim Page,totalPut:totalPut=Conn.Execute("Select count(1) From " & PostTable &" Where Verific=1 and TopicID=" & TopicID)(0)
			 If totalput Mod MaxPerPage = 0 Then
				Page=totalput\MaxPerPage
			 Else
				Page=totalput\MaxPerPage + 1
			 End If
			 Session("PopTips")=PopTips
			 Response.Write KS.GetClubShowUrlPage(TopicID,page)
			ElseIf KS.ChkClng(Session("TopicMusicReply"))=1 Then
			 Session("PopTips")=PopTips
			 Response.Write "gohome|"&KS.GetClubShowUrl(TopicID)
			Else
			 Dim UserXml,UN,LC
			 Dim Floor:Floor=Conn.Execute("Select Count(1) From " & PostTable &" Where TopicID=" & TopicID)(0)-1
			 Dim KesionClub:Set KesionClub=New ClubDisplayCls
			 Dim RSU:Set RSU=Conn.Execute("Select top 1 " & KesionClub.UserFields & " From KS_User Where UserName='" & UserName & "'")
			 If Not RSU.Eof Then Set UserXml=KS.RsToXml(RSU,"row","")
			 RSU.Close :Set RSU=Nothing
			 If IsObject(UserXML) Then set UN=UserXml.DocumentElement.SelectSingleNode("row[@username='" & TopicNode.SelectSingleNode("@username").text & "']") Else Set UN=Nothing
			 Set KesionClub.TopicNode=TopicNode
			 Set KesionClub.UN=UN
			 KesionClub.PostUserName=TopicNode.SelectSingleNode("@username").text
			 KesionClub.BSetting=BSetting
			 KesionClub.N=Floor+1
			 KesionClub.TopicID=TopicID
			 KesionClub.BoardID=BoardID
			 Set KesionClub.KSUser=KSUser
			 KesionClub.ReplayID=TopicNode.SelectSingleNode("@id").text
			 KesionClub.Immediate=false
			 KesionClub.Scan Application(KS.SiteSN&"LoopTemplate")
			 KS.Echo Escape(PopTips&"@@@@@"&KesionClub.Templates)
			 Set KesionClub=Nothing
			 KS.Die ""
			End If
	End Sub
		
	Sub SaveData()
			Dim O_LastPost,N_LastPost,O_LastPost_A
		    Dim SqlStr:SqlStr = "SELECT top 1 * From " & PostTable &" WHERE ID IS NULL" 
			Dim RSObj:Set RSObj=Server.CreateObject("Adodb.RecordSet")
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
			If KS.Setting(60)="1" and Check=false Then  
			RSObj("Verific")=0
			Else
			RSObj("Verific")=1
			End If
			RSObj("ParentId")=TopicID
			RSObj("DelTF")=0
			RSObj.Update
			RSObj.MoveLast
			Rid=RSObj("id")
			If KS.ChkClng(KS.S("toend"))=0 Then
			 Dim Xml: Set Xml=KS.RsToXml(RSObj,"","row")
			 Set TopicNode=Xml.DocumentElement.SelectSingleNode("row")
			End If
			RSObj.Close
			Set RSObj = Nothing
			'关联上传文件
			Call KS.FileAssociation(1036,RID,Content,0)
            If Not KS.IsNul(Session("UploadFileIDs")) Then 
				 Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & Rid &",classID=" & KS.ChkClng(Request.Form("BoardID")) & " Where ID In (" & KS.FilterIds(Session("UploadFileIDs")) & ")")
			End If
			
           If LoginTF=true Then
			  Conn.Execute("Update KS_User Set PostNum=PostNum+1 Where UserName='" & KSUser.UserName & "'")
			End If			

			Dim Subject:Subject=UnEscape(KS.S("Subject"))
			Conn.Execute("Update KS_GuestBook Set LastReplayTime=" & SqlNowString &",LastReplayUser='" & UserName &"',LastReplayUserID=" & UserID & ",TotalReplay=TotalReplay+1 where id=" & TopicID)
			
			N_LastPost=topicid & "$" & now & "$" & Replace(Subject,"$","") &"$" & UserName & "$" &UserID&"$$"
			
			If KS.ChkClng(BSetting(4))>0 and LoginTF=true Then
			     PopTips="<strong>积分" & KSUser.GetUserInfo("Score") &"+</strong>" & KS.ChkClng(BSetting(4))
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(4)),"系统","在论坛回复主题[" & Subject & "]所得!",0,0)
			End If
			If KS.ChkClng(BSetting(4))<0 and LoginTF=true Then
			    PopTips="<strong>积分" & KSUser.GetUserInfo("Score") &"-</strong>" & -KS.ChkClng(BSetting(4))
				Call KS.ScoreInOrOut(KSUser.UserName,2,-KS.ChkClng(BSetting(4)),"系统","在论坛回复主题[" & Subject & "]消费!",0,0)
			End If
			
             If LoginTF=true Then
			  If KS.ChkClng(BSetting(31))<>0 Then
			  if PopTips="" then
			   PopTips="<strong>威望" & KSUser.GetUserInfo("Prestige") &"+</strong>" & KS.ChkClng(BSetting(31))
			  Else
			   PopTips=PopTips & ",<strong>威望" & KSUser.GetUserInfo("Prestige") &"+</strong>" & KS.ChkClng(BSetting(31))
			  end if
			  If IsObject(Session(KS.SiteSN&"UserInfo")) Then Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row").SelectSingleNode("@prestige").Text=KS.ChkClng(KSUser.GetUserInfo("Prestige"))+KS.ChkClng(BSetting(30))

			  Conn.Execute("Update KS_User Set Prestige=Prestige+" & KS.ChkClng(BSetting(31)) & " Where UserName='" & KSUser.UserName &"'")
			  End If
			  Call KSUser.AddLog(KSUser.UserName,"在论坛回复了主题[<a href='{$GetSiteUrl}club/display.asp?id=" & TopicID & "' target='_blank'>" & subject &"</a>]",100)
			End If			
			
			'更新版面数据
			If BoardID<>0 Then
			  KS.LoadClubBoard()
			  O_LastPost=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text
			  
			  Conn.Execute("Update KS_GuestBoard set lastpost='" & N_LastPost & "',postnum=postnum+1 where id=" & BoardID)
				If KS.IsNul(O_LastPost) Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				Else
				 O_LastPost_A=Split(O_LastPost,"$")
				 Dim LastPostDate:LastPostDate=O_LastPost_A(1)
				 If Not IsDate(LastPostDate) Then LastPostDate=Now
				 If datediff("d",LastPostDate,Now())=0 Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=todaynum+1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text)+1
				 Else
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				 End If
				End If
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text)+1
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text=N_LastPost
			End If
			
			'更新今日发帖数等
			Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
            Set Doc=Nothing
		End sub
		
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
		
		function check()
	 	Dim KSLoginCls,Master
		Set KSLoginCls = New LoginCheckCls1
		If KSLoginCls.Check=true Then
		  check=true
		  Exit function
		else
		    master=LFCls.GetSingleFieldValue("select top 1 master from ks_guestboard where id=" & KS.ChkClng(FCls.RefreshFolderID))
			Dim KSUser:Set KSUser=New UserCls
			If Cbool(KSUser.UserLoginChecked)=false Then 
			  check=false
			  exit function
			else
			   check=KS.FoundInArr(master, KSUser.UserName, ",")
			End If
		end if
 End function	
End Class
%> 
