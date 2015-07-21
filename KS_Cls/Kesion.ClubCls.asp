<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Class ClubCls
        Private KS, KSR,ListStr,Node,BSetting,KSUser,GuestTitle,Master,MasterArr,FileContent,TopicID
		Private ListTemplate,pLoopTemplate,LoopTemplate,LoopList,boardid,parentId,PostBtnStr,TopXML,TopicXml,TopicNode
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno,Immediate,Templates
	    Private SqlStr,Doc,ListUrl,startime,LoginTF,CachePage,CacheTime
		Private Sub Class_Initialize()
		 CachePage=true   '��ҳ����,�����������������ϴ�ʱ,�������ó�true
		 CacheTime=0     '��ҳ����ʱ������,��λΪ����
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Immediate = true
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		%>
		<!--#include file="Kesion.IfCls.asp"-->
		<!--#include file="ClubFunction.asp"-->
		<%
		
		Public Sub Kesion()
		    startime=Timer()
			If KS.Setting(56)="0" Then Call KS.ShowTips("error","��վ�ѹر���̳����!") 
			If KS.Setting(59)="1" Then 
				 Dim P:P=KS.QueryParam("page")
				 If P="" Then
					response.Redirect(KS.Setting(3) & KS.Setting(66) & "/guestbook.asp")
				 Else
					response.Redirect(KS.Setting(3) & KS.Setting(66) & "/guestbook.asp?" & P)
				 End If
			End If

			FCls.RefreshType = "guestindex" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
			If Not KS.IsNul(Request.QueryString) Then 
				LoadClubBoardList
			Else
				LoadClubIndex
			End If
			GetClubPopLogin FileContent
			FileContent=KSR.ReplaceGeneralLabelContent(FileContent)
			FileContent=Replace(Replace(FileContent,"��#","{"),"#��","}")  '��ǩ�滻����
			FileContent=RexHtml_IF(FileContent)
			FileContent=Replace(FileContent,"{#ExecutTime}","ҳ��ִ��" & FormatNumber((timer()-startime),5,-1,0,-1) & "�� powered by <a href='http://www.kesion.com' target='_blank'>KesionCMS 7.0</a>")
			 KS.Echo FileContent
		End Sub
		
		'��ҳ
		Sub LoadClubIndex()
		    If KS.Setting(114)="" Then KS.Die "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!"
			FCls.RefreshFolderID = 0
			If KS.IsNUL(Application(KS.SiteSN&"ClubIndex")) or (isDate(Application(KS.SiteSn &"ClubIndexUpdateTime")) and DateDiff("n",Application(KS.SiteSn &"ClubIndexUpdateTime"),Now)>=CLng(CacheTime)) or CachePage=false Then
				Application(KS.SiteSn &"ClubIndexUpdateTime")=Now
				FileContent = KSR.LoadTemplate(KS.Setting(114))
				FileContent=KSR.ReplaceAllLabel(FileContent)
				FileContent=KSR.ReplaceLableFlag(FileContent)
				Application(KS.SiteSN&"ClubIndex")=FileContent
			 Else
			    FileContent = Application(KS.SiteSN&"ClubIndex")
			 End If
			 KS.LoadClubBoard : Call GetBoardList()
			 ListTemplate = LoopList
			set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
			If DateDiff("d",xmldate,now)=0 Then
					  If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) then
					   doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					   doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					  end if
			Else
					  GCls.Execute("Update KS_GuestBoard Set TodayNum=0")
				      Application(KS.SiteSN&"_ClubBoard")=empty	
					  doc.documentElement.attributes.getNamedItem("date").text=now
					  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  doc.documentElement.attributes.getNamedItem("todaynum").text=0
					  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			End If
	
			FileContent=Replace(FileContent,"{$TodayNum}",doc.documentElement.attributes.getNamedItem("todaynum").text)
			FileContent=Replace(FileContent,"{$YesterDayNum}",doc.documentElement.attributes.getNamedItem("yesterdaynum").text)
			FileContent=Replace(FileContent,"{$MaxDayNum}",doc.documentElement.attributes.getNamedItem("maxdaynum").text)
			FileContent=Replace(FileContent,"{$TopicNum}",doc.documentElement.attributes.getNamedItem("topicnum").text)
			FileContent=Replace(FileContent,"{$ReplayNum}",doc.documentElement.attributes.getNamedItem("postnum").text)
			FileContent=Replace(FileContent,"{$UserNum}",doc.documentElement.attributes.getNamedItem("totalusernum").text)
			FileContent=Replace(FileContent,"{$NewUser}",doc.documentElement.attributes.getNamedItem("newreguser").text)
			FileContent=Replace(FileContent,"{$MaxOnline}",doc.documentElement.attributes.getNamedItem("maxonline").text)
			FileContent=Replace(FileContent,"{$MaxOnlineDate}",doc.documentElement.attributes.getNamedItem("maxonlinedate").text)
			PostBtnStr="<a href=""javascript:Posted()""><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_post.png"" align=""absmiddle"" alt=""����""></a>"
			FileContent=Replace(FileContent,"{$PostButtonAction}",PostBtnStr)
			FileContent=Replace(FileContent,"{$GuestTitle}",KS.Setting(61))
			FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
		End Sub
	    '����
		Sub LoadClubBoardList()
		   If KS.Setting(172)="" Then KS.Die "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!"
		   FileContent = KSR.LoadTemplate(KS.Setting(172))
		   If Not KS.IsNul(KS.Setting(69)) and Request.QueryString<>"" Then
					  Dim QueryStr:QueryStr=Request.QueryString
					  Dim QArr:QArr=Split(Split(QueryStr,".")(0),"-")
					  If Ubound(Qarr)>=1 Then
					   BoardID=KS.ChkClng(Qarr(1))
					  Else
					   BoardID=KS.ChkClng(KS.S("BoardID"))
					  End If
					  If Ubound(QArr)>=2 Then  
					   CurrentPage = KS.ChkClng(Qarr(2))
					  Else
					   CurrentPage = KS.ChkClng(Request("page")) 
					  End If
			Else
					  BoardID=KS.ChkClng(KS.S("BoardID"))
					  CurrentPage = KS.ChkClng(Request("page")) 
			End If
			FCls.RefreshFolderID = BoardID '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
		    KS.LoadClubBoard
			Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			If Node Is Nothing Then KS.Die "�Ƿ�����!"
			BSetting=Node.SelectSingleNode("@settings").text
			ParentId=KS.ChkClng(Node.SelectSingleNode("@parentid").text)
			FileContent=Replace(FileContent,"{$BoardName}",Node.SelectSingleNode("@boardname").text)
			FileContent=Replace(FileContent,"{$GetBoardUrl}",KS.GetClubListUrl(boardid))
			Master=Node.SelectSingleNode("@master").text
			
             BSetting=BSetting&"$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
			 BSetting=Split(BSetting,"$")
			 If CurrentPage<=0 Then CurrentPage=1
			 MaxPerPage=KS.ChkClng(BSetting(20)) : If MaxPerPage=0 Then MaxPerPage=KS.ChkClng(KS.Setting(51))

			 If Not KS.IsNul(KS.Setting(69)) Then
			  ListUrl="http://" & KS.Setting(69) & "/"
			 Else
			  ListUrl=KS.GetDomain & KS.Setting(66) &"/"
			 End If
				
			If  BSetting(0)="0" And KS.C("UserName")="" Then
					ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error1")
					GuestTitle="��Ȩ����"
			ElseIf parentid<>0 or KS.S("Istop")="1" or KS.S("IsBest")="1" then
				       LoginTF=KSUser.UserLoginChecked
					   Dim GroupPurview:GroupPurview= True : If Not KS.IsNul(BSetting(1)) and (KS.FoundInArr(Replace(BSetting(1)," ",""),KSUser.GroupID,",")=false Or LoginTF=false) Then GroupPurview=false
					   Dim UserPurview:UserPurview=True : If Not KS.IsNul(BSetting(10)) and (KS.FoundInArr(BSetting(10),KSUser.UserName,",")=false or LoginTF=false) Then UserPurview=false
					   If KSUser.GetUserInfo("ClubSpecialPower")="1" Then UserPurview=true:GroupPurview=True
					   Dim ScorePurview:ScorePurview=KS.ChkClng(BSetting(11))
					   Dim MoneyPurview:MoneyPurview=KS.ChkClng(BSetting(12))
					   
					   If ((GroupPurview=false and Not KS.IsNul(BSetting(10))) or (UserPurview=false)) and boardid<>0 Then
					    ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error2")
						GuestTitle="��Ȩ����"
					   ElseIf KS.ChkClng(KSUser.GetUserInfo("Score"))<ScorePurView And ScorePurView>0 Then
					    ListTemplate=Replace(Replace(LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error5"),"{$Tips}","����<span>" &ScorePurView&"</span>��"),"{$CurrTips}","����<span>" & KSUser.GetUserInfo("Score") & "</span>��")
						
						GuestTitle="��Ȩ����"
					   ElseIf KS.ChkClng(KSUser.GetUserInfo("Money"))<MoneyPurview And MoneyPurview>0 Then
					    ListTemplate=Replace(Replace(LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error5"),"{$Tips}","�ʽ�<span>" &formatnumber(MoneyPurview,2,-1,-1)&"</span>Ԫ"),"{$CurrTips}","�ʽ�<span>" & formatnumber(KSUser.GetUserInfo("money"),2,-1,-1) & "</span>Ԫ")
						GuestTitle="��Ȩ����"
					   Else
						   
						   if boardid<>0  Then
						    GuestTitle=KS.LoseHtml(Node.SelectSingleNode("@boardname").text)
						   else
							if KS.S("Istop")="1" then GuestTitle="�ö�����" Else GuestTitle="��������"
						   end if
							PostBtnStr="<span style=""position:relative;"" onmouseover=""$('#postlist').show()"" onmouseout=""$('#postlist').hide()""><a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_post.png""></a><div id=""postlist"" class=""submenu noli"">"
							PostBtnStr=PostBtnStr&"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/new_post.gif"" align=""absmiddle""/> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """>��������</a></dl>"
							PostBtnStr=PostBtnStr &"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/vote.gif"" align=""absmiddle""> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & BoardID&"&posttype=1"">����ͶƱ</a></dl>"
							PostBtnStr=PostBtnStr &"</div></span>"
							Call GetLoopList()
							GetClubPopLogin FileContent
							FileContent=Replace(FileContent,"{$PostButtonAction}",PostBtnStr)						   
				            FileContent=Replace(FileContent,"{$GuestTitle}",GuestTitle)
						    FileContent=RexHtml_IF(FileContent) '�ȹ������õı�ǩ,���ٱ�ǩ����
						    FileContent=KSR.KSLabelReplaceAll(FileContent)
						  ' ks.die filecontent
						   Scan FileContent
						   ks.die ""
						   ListTemplate = Replace(ListTemplate,"[loop]" & LoopTemplate &"[/loop]",LoopList)
					 End If
				Else
				 KS.LoadClubBoard : Call GetBoardList()
				 ListTemplate=LoopList
				END IF
				FileContent=RexHtml_IF(FileContent) '�б�ҳ�ȹ���������ǩ,���ٱ�ǩ����

				FileContent=Replace(FileContent,"{$GuestTitle}",GuestTitle)
				FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
				FileContent=Replace(Replace(Replace(Replace(Replace(FileContent,"{$ShowManageCheckBox}",""),"{$Img}",""),"{$PageList}",""),"{$Jing}",""),"{$Status}","") '�滻�����ñ�ǩ,�ӿ����
				FileContent=KSR.ReplaceAllLabel(FileContent)
				FileContent=KSR.ReplaceLableFlag(FileContent)
                FileContent=Replace(FileContent,"{$BoardID}",boardid)
		End Sub	
		
		Function Parse(sTemplate, iPosBegin)
			Dim iPosCur, sToken, sValue, sTemp
			iPosCur        = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			iPosCur       = InStr(sTemp, ".")
			if iPosCur>1 Then
			sToken        = Left(sTemp, iPosCur-1)
			End If
			sValue        = Mid(sTemp, iPosCur+1) 
		
			Select Case lcase(sValue)
				Case "begin"
					sTemp            = "{@" & ( sToken & ".end}" )
					iPosCur            = InStr(iPosBegin, sTemplate, sTemp)
					ParseArea      sToken, Mid(sTemplate, iPosBegin, iPosCur-iPosBegin)
					iPosBegin        = iPosCur+Len(sTemp)
				case "boardid" echo boardid
				case "boardname" echo Node.SelectSingleNode("@boardname").text
				case "boardintro" echo Node.SelectSingleNode("@note").text

				case "master"
				    If KS.IsNul(Master) Then 
					  Echo "<a href='#'>���ް���</a>"
					Else
					 If Not IsObject(Application(KS.SiteSN &"Master"&BoardID)) Then
					   Call LoadMasterUserID(BoardID,Master)
					 End If
					 Dim MyMaster:MyMaster=Application(KS.SiteSN &"Master"&BoardID)
					 If Not KS.IsNul(MyMaster) Then
						 MasterArr=Split(MyMaster,"@") 
						 For I=0 To Ubound(MasterArr)
						   If I=0 Then echo "<a href='" & KS.GetSpaceUrl(Split(MasterArr(i),"|")(0)) & "' target='_blank'>" & Split(MasterArr(i),"|")(1) & "</a>" Else echo "," & "<a href='" & KS.GetSpaceUrl(Split(MasterArr(i),"|")(0)) & "' target='_blank'>" & Split(MasterArr(i),"|")(1) & "</a>"
						 Next
					 End If
					End If
			   case "boardrules" echo Node.SelectSingleNode("@boardrules").text
			   case "executtime" echo "ҳ��ִ��" & FormatNumber((timer()-startime),5,-1,0,-1) & "�� powered by <a href='http://www.kesion.com' target='_blank'>KesionCMS 7.0</a>"
			   case "showpage"
			    If Not KS.IsNul(Request("a")) or Not KS.IsNul(Request("c")) or Not KS.IsNul(Request("d"))  or Not KS.IsNul(Request("o")) or Not KS.IsNul(Request("isbest")) or Not KS.IsNul(Request("istop")) Then
				   echo KS.ShowPage(TotalPut,MaxPerPage,"",CurrentPage,false,false)
				Else
				   echo KS.GetClubPageList(MaxPerPage,CurrentPage,TotalPut,BoardID,Gcls.ClubPreList)
				End If
				Case Else
					ParseNode sToken, sValue
		   End Select
		   Parse    = iPosBegin
		End Function
        Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
			  Case "toploop"
			    LoadTopTopic
			    If CurrentPage=1 And IsObject(Application(KS.SiteSN &"TopXML")) Then
				  For Each TopicNode In Application(KS.SiteSN &"TopXML").DocumentElement.SelectNodes("row[@boardid=" & Boardid&" or @istop=2]")
				     TopicID=TopicNode.SelectSingleNode("@id").text
					 scan sTemplate
				  Next
				  echo "<table border=""0"" style=""margin:0px auto;width:98%"" align=""center"" class=""topiclist"" cellpadding=""0"" cellspacing=""0""><tr><td style=""background:#FAFDFF;height:25px;padding-left:15px"">===��ͨ����===</td></tr></table>"
				End If
			  Case "loop"
			    If IsObject(TopicXML) Then
				  For Each TopicNode In TopicXML.DocumentElement.SelectNodes("row")
				     TopicID=TopicNode.SelectSingleNode("@id").text
					 scan sTemplate
				  Next
				End If
				
			End Select
		End Sub
		Sub ParseNode(sTokenType, sTokenName)
					Select Case lcase(sTokenType)
					    case "item"
						  select case lcase(sTokenName)
						    case "ico" 
							  dim IcoUrl,TitleTips
							  If KS.ChkClng(TopicNode.SelectSingleNode("@posttype").text)=1 Then
			                   IcoUrl="vote.gif" : TitleTips="ͶƱ����"
							  ElseIf cint(TopicNode.SelectSingleNode("@istop").text)=1 Then
							   IcoUrl="top.gif" : TitleTips="�������ö�"
							  ElseIf cint(TopicNode.SelectSingleNode("@istop").text)=2 Then
							   IcoUrl="ztop.gif": TitleTips="���ö�"
							  ElseIf cint(TopicNode.SelectSingleNode("@verific").text)=2 Then
							   IcoUrl="lock.gif": TitleTips="��������"
							  ElseIf KS.ChkClng(TopicNode.SelectSingleNode("@hits").text)>KS.ChkClng(BSetting(27)) and KS.ChkClng(TopicNode.SelectSingleNode("@totalreplay").text)>KS.ChkClng(BSetting(28)) Then
							   IcoUrl="hot.gif": TitleTips="��������"
							  Else
							   IcoUrl="common.gif": TitleTips="��ͨ����"
							  End If
							  echo "<a href='" & KS.GetClubShowUrl(TopicID) &"' target='_blank'><img border='0' src='" & KS.Setting(3) & KS.Setting(66) & "/images/" & IcoUrl & "' title='" & TitleTips & "'></a>"
							case "author" 
							  Dim PostUser:PostUser=TopicNode.SelectSingleNode("@username").text
							  If KS.IsNul(PostUser) Then
							   echo "<a href=""#"" class=""author"" target=""_blank"">�ο�</a>"
							  Else
							   echo "<a href=""" & KS.GetSpaceUrl(TopicNode.SelectSingleNode("@userid").text) & """ class=""author"" target=""_blank"">" & PostUser& "</a>"
							  End If
							case "pubtime" echo KS.GetTimeFormat(TopicNode.SelectSingleNode("@addtime").text)
							case "replaytimes" echo TopicNode.SelectSingleNode("@totalreplay").text
							case "hits" echo TopicNode.SelectSingleNode("@hits").text 
							case "lastreplayuser"
							  dim LastReplayUser:LastReplayUser=TopicNode.SelectSingleNode("@lastreplayuser").text
							  If KS.IsNul(LastReplayUser) Then
							   echo "<a href=""#"" target=""_blank"">�ο�</a>"
							  Else
							   echo "<a href=""" & KS.GetSpaceUrl(TopicNode.SelectSingleNode("@lastreplayuserid").text) & """ class=""author"" target=""_blank"">" & LastReplayUser& "</a>"
							  End If
							case "lastreplaytime" echo KS.GetTimeFormat1(TopicNode.SelectSingleNode("@lastreplaytime").text,true)
							case "subjectlist"
							   If KS.S("A")="m" Then echo "<input type='checkbox' name='m' onclick=""showmanage(this.checked,this.value,'" & KS.Setting(66) & "'," & BoardID & ")"" value='" & TopicID & "'/>"
							   If KS.ChkClng(BSetting(25))>0 and isobject(Application(KS.SiteSN&"_ClubBoardCategory")) Then
								Dim CategoryNode,CategoryId,categoryName,categoryIco
								CategoryId=TopicNode.SelectSingleNode("@categoryid").text
								Set CategoryNode=Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectSingleNode("row[@categoryid=" & CategoryId&"]")
								If Not CategoryNode Is Nothing Then
								categoryname=CategoryNode.SelectSingleNode("@categoryname").text : If Instr(categoryname,"[")=0 and categoryname<>"" Then categoryname="[" & categoryname & "]"
								categoryIco=CategoryNode.SelectSingleNode("@ico").text
									If KS.ChkClng(BSetting(25))=2 Then
									echo " <a href=""" & ListUrl & "?boardid=" & boardid& "&c=" &CategoryId&"""><Img src='" & categoryIco & "' alt='" &CategoryName & "' border='0' align='absmiddle'/></a>"
									Else
									echo "<a href=""" & ListUrl & "?boardid=" & boardid& "&c=" &CategoryId&""">" & CategoryName &"</a>"
									End If
								End If
							  End If
							  echo "<a href=""" & KS.GetClubShowUrl(TopicID) & """>" & replace(replace(TopicNode.SelectSingleNode("@subject").text,"��#","{"),"#��","}") & "</a>"
							  if cint(TopicNode.SelectSingleNode("@verific").text)=0 Then
							   echo " <span style='color:red'>[δ��]</span>"
							  ElseIf cint(TopicNode.SelectSingleNode("@verific").text)=2 Then
							   echo " <span style='color:green'>[����]</span>"
							  End If
							  If cint(TopicNode.SelectSingleNode("@isbest").text)=1 Then echo "<Img src='" & KS.Setting(3) & KS.Setting(66) & "/images/jing.gif' border='0' alt='��������' align='absmiddle'/> "
							  Dim AnnexExt,TotalReplay,MaxPage,pages,K
							  AnnexExt=TopicNode.SelectSingleNode("@annexext").text
							  If Not KS.IsNul(AnnexExt) Then
			                   echo " <Img src='" & KS.Setting(3) & "editor/ksplus/fileicon/" & AnnexExt &".gif' alt='" & AnnexExt & "����' border='0' align='absmiddle'/>"
			                  Else
								  If KS.ChkClng(TopicNode.SelectSingleNode("@ispic").text)=1 Then
									echo " <Img src='" & KS.Setting(3) & KS.Setting(66) & "/images/image_s.gif' alt='GifͼƬ����' border='0' align='absmiddle'/>"
								  ElseIf KS.ChkClng(TopicNode.SelectSingleNode("@ispic").text)=2 Then
									echo " <Img src='" & KS.Setting(3) & KS.Setting(66) & "/images/image_s.gif' alt='JPGͼƬ����' border='0' align='absmiddle'/>"
								  End If
							  End If
							  
							  '����Ŀ�߷�ҳ
							  TotalReplay=KS.ChkClng(TopicNode.SelectSingleNode("@totalreplay").text)
							  If TotalReplay<>0 Then
							     MaxPage=KS.ChkClng(BSetting(21)) : If MaxPage=0 Then MaxPage=10
								 If TotalReplay Mod MaxPage = 0 Then
										Pages=TotalReplay\MaxPage
								 Else
										Pages=TotalReplay\MaxPage + 1
								 End If
							   If Pages>1 Then
									    echo "<span class=""topic-pages""><img src='" &KS.Setting(3) & KS.Setting(66) & "/images/multipage.gif' title='��ҳ'/>"
										if pages>5 then
										   echo " <a href='" & KS.GetClubShowUrlPage(TopicID,1) & "'>1</a>"
										   echo " <a href='" & KS.GetClubShowUrlPage(TopicID,2) & "'>2</a>"
										   echo "..."
										   echo " <a href='" & KS.GetClubShowUrlPage(TopicID,pages-3) & "'>" & pages-3 &"</a>"
										   echo " <a href='" & KS.GetClubShowUrlPage(TopicID,pages-2) & "'>" & pages-2 &"</a>"
										   echo " <a href='" & KS.GetClubShowUrlPage(TopicID,pages-1) & "'>" & pages-1 &"</a>"
										   echo " <a href='" & KS.GetClubShowUrlPage(TopicID,pages) & "'>" & pages &"</a>"
										Else
										   For k=1 to Pages
											 echo " <a href='" & KS.GetClubShowUrlPage(TopicID,k) & "'>"&k&"</a>"
										   Next
										End If
								   echo "</span>"
								End if
							  End If
							  If KS.ChkClng(BSetting(42))<>0 Then
							   If DateDiff("h",TopicNode.SelectSingleNode("@lastreplaytime").text,now)<=KS.ChkClng(BSetting(42)) Then
							  echo " <img src='" &KS.Setting(3) & KS.Setting(66) & "/images/new.gif' />"
							   End If
							  End If
						  end select
					end select
		End Sub
		'�г�����
		Sub GetBoardList()
		  Dim LC,PNode,Node,Xml,Str,TStr,Bparam,LastPost,LastPost_A
          Set Xml=Application(KS.SiteSN&"_ClubBoard")
		  If parentid=0 and boardid<>0 Then BParam="id=" & boardid Else BParam="parentid=0"
		  If IsObject(xml) Then
		       PLoopTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","boardclass")
		       LoopTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","board")
			   For Each Pnode In Xml.DocumentElement.SelectNodes("row[@" & BParam & "]")
					 LC=PLoopTemplate
					 GuestTitle=PNode.SelectSingleNode("@boardname").text
					 LC=Replace(LC,"{$BoardUrl}",KS.GetClubListUrl(PNode.SelectSingleNode("@id").text))
					 LC=replace(LC,"{$BoardID}",PNode.SelectSingleNode("@id").text)
					 LC=replace(LC,"{$BoardName}",PNode.SelectSingleNode("@boardname").text)
					 LC=replace(LC,"{$Intro}",PNode.SelectSingleNode("@note").text)
					 If KS.IsNul(PNode.SelectSingleNode("@master").text) then
					 LC=replace(LC,"{$Master}","���ް���")
					 else
					 LC=replace(LC,"{$Master}",PNode.SelectSingleNode("@master").text)
					 end if
					 LC=replace(LC,"{$TotalSubject}",PNode.SelectSingleNode("@topicnum").text)
					 LC=replace(LC,"{$TotalReply}",PNode.SelectSingleNode("@postnum").text)
					 LC=replace(LC,"{$TodayNum}",PNode.SelectSingleNode("@todaynum").text)
                     
					 tstr=""
					 
				   For Each Node In Xml.DocumentElement.SelectNodes("row[@parentid=" & Pnode.SelectSingleNode("@id").text & "]")
					 str=LoopTemplate
					 str=Replace(str,"{$BoardUrl}",KS.GetClubListUrl(Node.SelectSingleNode("@id").text))
					 str=replace(str,"{$BoardID}",Node.SelectSingleNode("@id").text)
					 str=replace(str,"{$BoardName}",Node.SelectSingleNode("@boardname").text)
					 str=replace(str,"{$Intro}",Node.SelectSingleNode("@note").text)
					 If KS.IsNul(Node.SelectSingleNode("@master").text) then
					 str=replace(str,"{$Master}","���ް���")
					 else
					 str=replace(str,"{$Master}",Node.SelectSingleNode("@master").text)
					 end if
					 
					 LastPost=Node.SelectSingleNode("@lastpost").text
					 If KS.IsNul(LastPost) Then
					  str=replace(str,"{$NewTopic}","��")
					  str=replace(str,"{$LastPostUser}","��")
					  str=replace(str,"{$LastPostTime}","-")
					 Else
					  LastPost_A=Split(LastPost,"$")
					  If LastPost_A(0)="0" or LastPost_A(2)="��" then
					  str=replace(str,"{$NewTopic}","��")
					  str=replace(str,"{$LastPostUser}","��")
					  str=replace(str,"{$LastPostTime}","-")
					  else
					  str=replace(str,"{$NewTopic}","<a href='" & KS.GetClubShowUrl(LastPost_A(0)) & "'>" & KS.gottopic(KS.LoseHtml(Replace(LastPost_A(2),"{","��#")),38) & "</a>")
					  str=replace(str,"{$LastPostUser}","<a href='" & KS.GetSpaceUrl(KS.ChkClng(LastPost_A(4))) &"' target='_blank'>" &LastPost_A(3) & "</a>")
					  str=replace(str,"{$LastPostTime}",KS.GetTimeFormat1(LastPost_A(1),true))

					  end if
					 End If

					 str=replace(str,"{$TotalSubject}",Node.SelectSingleNode("@topicnum").text)
					 str=replace(str,"{$TotalReply}",Node.SelectSingleNode("@postnum").text)
					 str=replace(str,"{$TodayNum}",Node.SelectSingleNode("@todaynum").text)
					 TStr=TStr&str
				  Next
					LC=Replace(LC,"<!--boardlist-->",tstr)
				  LoopList=LoopList & LC
			 Next
		  End If
		End Sub
		'�г�����
		Sub GetLoopList()
		    Dim ListType,Param
			Dim OrderArr:OrderArr=Array("Ĭ������|0|0","����ID�š�|1|0","����ID�š�|1|1","� �� ����|2|0","�ظ�ʱ���|0|0","�ظ�ʱ���|0|1","� �� ����|2|0","� �� ����|2|1","�� �� ����|3|0","�� �� ����|3|1")
			Dim DateArr:DateArr=Array("ȫ��ʱ��|0","һ��|1","����|3","һ����|7","һ������|30","��������|90","������|180","һ����|365")
		    If KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" Or (KSUser.GetUserInfo("ClubSpecialPower")="3" and KS.FoundInArr(Master,KSUser.UserName,",")=true) Then 
			 Param=" Where deltf=0"	
			 if KS.S("A")="m" then
			   FileContent=Replace(FileContent,"{$ShowManageButton}","<a href=""" & KS.GetClubListUrl(boardid) & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_manage.png""></a>")
			 else
			   FileContent=Replace(FileContent,"{$ShowManageButton}","<a href=""" & ListUrl & "?a=m&boardid=" & BoardID & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_manage.png""></a>")
			 end if
			Else  
			 Param=" Where deltf=0 and verific<>0"
			 if KS.S("a")="m" then
			   KS.Die "<script>alert('��û�й����Ȩ��,�벻Ҫ�Ƿ�����!');history.back(-1);</script>"
			 end if
			End If
			
			ListType="<li>���⣺</li>"
			if request.querystring.count=1 then
			 ListType=ListType & "<li class=""current""><a href='" & ListUrl &"?boardid=" & boardid & "'>ȫ��</a></li>"
			else
			 ListType=ListType & "<li><a href='" & ListUrl &"?boardid=" & boardid & "'>ȫ��</a></li>"
			end if
			If KS.ChkClng(KS.S("Istop"))=1 Then 
			 Param=Param & " and istop<>0"
			 ListType=ListType & "<li class=""current""><a href='" & ListUrl &"?boardid=" & boardid & "&istop=1'>�ö�</a></li>"
			Else
			 ListType=ListType & "<li><a href='" & ListUrl &"?boardid=" & boardid & "&istop=1'>�ö�</a></li>"
			End If
			If KS.ChkClng(KS.S("IsBest"))=1 Then 
			 Param=Param & " and isbest=1"
			 ListType=ListType & "<li class=""current""><a href='" & ListUrl &"?boardid=" & boardid & "&isbest=1'>����</a></li>"
			Else
			 ListType=ListType & "<li><a href='" & ListUrl &"?boardid=" & boardid & "&isbest=1'>����</a></li>"
			End If
			ListType=ListType & "&nbsp;&nbsp;<li>| &nbsp;&nbsp; </li>"
            
			Dim D:D=KS.ChkClng(KS.S("D"))
			Dim O:O=KS.ChkClng(KS.S("O"))
			Dim C:C=KS.ChkClng(KS.S("C"))
			'��ʱ��鿴
			ListType=ListType & "<li style=""position:relative;_padding-top:6px"" onmouseover=""$('#datelist').show()"" onmouseout=""$('#datelist').hide()"">" & vbcrlf
			if d<=Ubound(DateArr) Then
			  ListType=ListType & "<a href=""#"">" & split(DateArr(d),"|")(0) & " <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			  If D<>0 Then Param=Param & " and datediff(" & DataPart_D & ",AddTime," & SQLNowString &")<" & split(DateArr(d),"|")(1)
			Else
			ListType=ListType & "<a href=""#"">ȫ��ʱ�� <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			End If
			ListType=ListType & "<div id=""datelist"" class=""submenu"" style=""left:0px;"">" & vbcrlf
			For I=0 To Ubound(DateArr)
			  ListType=ListType & "<dl><a href=""" & ListUrl & "?boardid=" & boardid & "&d=" & I & """>" & Split(DateArr(i),"|")(0) &"</a></dl>"
			Next
			ListType=ListType & "</div></li>" & vbcrlf
			ListType=ListType & "&nbsp;&nbsp;<li>| &nbsp;&nbsp; </li>"
			'����ʽ
			ListType=ListType & "<li style=""position:relative;_padding-top:6px"" onmouseover=""$('#orderlist').show()"" onmouseout=""$('#orderlist').hide()"">" & vbcrlf
			if O<=Ubound(OrderArr) Then
			  ListType=ListType & "<a href=""#"">" & split(OrderArr(o),"|")(0) & " <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			Else
			ListType=ListType & "<a href=""#"">Ĭ������ <img src=""" & KS.Setting(3) & KS.Setting(66) &"/images/arw_d2.gif"" align=""absmiddle""/></a>" & vbcrlf
			End If
			ListType=ListType & "<div id=""orderlist"" class=""submenu"" style=""left:0px;"">" & vbcrlf
			For I=0 To Ubound(OrderArr)
			  ListType=ListType & "<dl><a href=""" & ListUrl & "?boardid=" & boardid & "&o=" & I & """>" & Split(OrderArr(i),"|")(0) &"</a></dl>"
			Next
			ListType=ListType & "</div></li>" & vbcrlf
			
		    FileContent=Replace(FileContent,"{$ListType}",ListType)
			
			'�������
			If BSetting(23)="1" And BSetting(26)="1" Then
			  KS.LoadClubBoardCategory
			  Dim CategoryNode,CategoryXML,CategoryStr,categoryImg
			  Set CategoryXML=Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectNodes("row[@boardid=" &BoardID &"]")
			  If CategoryXML.length>0 Then 
				  CategoryStr="<p class=""boardcategory cl"">" & vbcrlf
				  If C=0 Then
				   CategoryStr=CategoryStr & "<strong class=""otp brw"">ȫ��</strong>" &vbcrlf
				  Else
				   Param=Param & " and categoryId=" & KS.ChkClng(KS.S("C"))
				   CategoryStr=CategoryStr & "<a href='" & KS.GetClubListUrl(boardid) & "' class='brw'>ȫ��</a>" &vbcrlf
				  End If
				  For Each CategoryNode In CategoryXML
				   If CategoryNode.SelectSingleNode("@ico").text<>"" Then
				   categoryImg="<img class=""vm"" src=""" & CategoryNode.SelectSingleNode("@ico").text &""" /> "
				   Else
				   categoryImg=""
				   End If
				   If trim(C)=trim(CategoryNode.SelectSingleNode("@categoryid").text) Then
				  CategoryStr=CategoryStr & "<strong class=""otp brw"">" & categoryImg & CategoryNode.SelectSingleNode("@categoryname").text & "</strong>" &vbcrlf
				   Else
				     CategoryStr=CategoryStr & "<a href=""" & ListUrl & "?boardid=" & boardid & "&c=" &CategoryNode.SelectSingleNode("@categoryid").text &""" class=""brw"">" & categoryImg & CategoryNode.SelectSingleNode("@categoryname").text & "</a>"
				   End If

				  Next
				  CategoryStr=CategoryStr &"</p>"
			  End If
		      FileContent=Replace(FileContent,"{$BoardCategory}",CategoryStr)
			  
			End If

		  If BoardID<>0 Then Param=Param &" and boardid=" & boardid
          
		  Dim RS,ListTopicFields
		  ListTopicFields="ID,UserName,UserID,Subject,AddTime,Verific,LastReplayUser,LastReplayUserID,LastReplayTime,TotalReplay,BoardID,Hits,IsPic,IsTop,IsBest,PostType,AnnexExt,CategoryId" rem �����б��õ����ֶ�

		  
		  Param=Param & " and istop=0"
		  If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ClubLists"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
				Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
				Cmd.Parameters.Append cmd.CreateParameter("@inConditions",200,1,220)
				Cmd.Parameters.Append cmd.CreateParameter("@ListFields",200,1,220)
				Cmd.Parameters.Append cmd.CreateParameter("@inOrder",3)
				Cmd.Parameters.Append cmd.CreateParameter("@inSort",3)
				Cmd("@pagenow")=CurrentPage
				Cmd("@pagesize")=MaxPerPage
				Cmd("@inConditions")=param
				Cmd("@ListFields")=ListTopicFields
				Cmd("@inOrder")=split(OrderArr(o),"|")(1)
				Cmd("@inSort")=split(OrderArr(o),"|")(2)
				Set Rs=Cmd.Execute
				totalPut=GCls.Execute("Select Count(1) From KS_GuestBook " & Param)(0)
				If Not RS.Eof Then Set TopicXML=KS.RsToXml(RS,"row","")
		  Else
			 Dim OrderField,SortStr
			 Select Case split(OrderArr(o),"|")(1)
			  case 1 OrderField="Id"
			  case 2 OrderField="hits"
			  case 3 OrderField="TotalReplay"
			  case else OrderField="LastReplayTime"
             End Select
			 If split(OrderArr(o),"|")(2)=0 Then SortStr="Desc" Else SortStr="ASC"
	 
			 If CurrentPage=1 Then
			  SqlStr = "SELECT Top " & MaxPerPage & " " & ListTopicFields & " From KS_GuestBook " & Param &" ORDER BY IsTop Desc," & OrderField & " " & SortStr 
			 Else
			  SqlStr = "SELECT " & ListTopicFields & " From KS_GuestBook " & Param &" ORDER BY " & OrderField & " " & SortStr  
			 End If
			 Set RS=GCls.Execute(sqlstr)
			 IF RS.Eof And RS.Bof Then
				  totalput=0
				  LoopList = "<tr><td colspan=5>�˰���û��" & KS.Setting(62) & "!</td></tr>"
				  exit sub
			  Else
								TotalPut=GCls.Execute("Select Count(1) From KS_GuestBook " & Param)(0)
								If CurrentPage < 1 Then CurrentPage = 1
			
								If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								Set TopicXML=KS.ArrayToXml(RS.GetRows(MaxPerPage),rs,"row","")
				End IF
		 End If	
		   RS.Close:Set RS=Nothing
		End Sub
		
End Class


Class ClubDisplayCls
        Private KS, KSR,ListStr,ID,Node,managestr,TotalReplay,TreplayNum,PostTable
		Private ListTemplate,LoopTemplate,LoopList,FileContent,RST,master,PostType,CheckIsMaster
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno,ShowScore,IsBest,IsTop,DelTF,Verific,Subject,Hits
	    Private SqlStr,GuestTitle,AllowShow,CategoryID,CategoryNode,categoryname,startime
		Public UserFields,PostUserName,PostUserID,BSetting,N,KSUser,LoginTF,TopicID,BoardID
		Public ReplayID,XML,TopicNode,UserXML,CommentXML,Un,Immediate,Templates
		Private LC,UserNames,PIDS,RS
		Private re
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		 Immediate = true
		 UserFields="UserID,UserName,UserFace,Sign,Sex,Score,Prestige,BestTopicNum,LoginTimes,RegDate,email,qq,postNum,LastLoginTime,ClubGradeID,IsOnline,LockOnClub"
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		%>
		<!--#include file="Kesion.IfCls.asp"-->
		<!--#include file="ClubFunction.asp"-->
		<%
		Function Parse(sTemplate, iPosBegin)
			Dim iPosCur, sToken, sValue, sTemp
			iPosCur        = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			iPosCur       = InStr(sTemp, ".")
			if iPosCur>1 Then
			sToken        = Left(sTemp, iPosCur-1)
			End If
			sValue        = Mid(sTemp, iPosCur+1) 
		
			Select Case lcase(sValue)
				Case "begin"
					sTemp            = "{@" & ( sToken & ".end}" )
					iPosCur            = InStr(iPosBegin, sTemplate, sTemp)
					ParseArea      sToken, Mid(sTemplate, iPosBegin, iPosCur-iPosBegin)
					iPosBegin        = iPosCur+Len(sTemp)
				case "subject" echo Replace(Replace(subject,"��#","{"),"#��","}")
				case "subjectnohtml" echo KS.CheckXSS(KS.LoseHtml(Replace(Replace(subject,"��#","{"),"#��","}")))
				case "description" 
				 If IsObject(Xml) Then
				 Set TopicNode=Xml.DocumentElement.SelectSingleNode("row[@parentid='0']/@content")
				  If Not TopicNode Is   Nothing Then  echo KS.Gottopic(KS.LoseHtml(Replace(Ubbcode(topicnode.text,0),chr(10),"")),150)
				 End If
				case "hits" echo hits
				case "totalreplay" 
				 If KS.ChkClng(totalreplay)>0 Then echo totalreplay-1 Else Echo 0
				case "guesttitle" echo guesttitle
				case "topicid" echo TopicID
				case "boardid" echo boardid
				case "boardurl" KS.GetClubListUrl(BoardID)
				case "posttable" echo PostTable
		        
				
				case "executtime" echo "ҳ��ִ��" & FormatNumber((timer()-startime),5,-1,0,-1) & "�� powered by <a href='http://www.kesion.com' target='_blank'>KesionCMS 7.0</a>"
				case "boardcategory"
						   If CategoryID<>0 Then
							   KS.LoadClubBoardCategory
							   Set CategoryNode=Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectSingleNode("row[@categoryid=" &CategoryID &"]")
							   If Not CategoryNode Is Nothing Then
							   categoryname=CategoryNode.SelectSingleNode("@categoryname").text : If Instr(categoryname,"[")=0 Then categoryname="[" & categoryname & "]"
							   echo categoryname
							   End If
							   Set CategoryNode=Nothing
						   End If
				case "managemenu"
					  echo "<li class=""backlist""><a href=""" & KS.GetClubListUrl(boardid) & """> << �����б�</a></li>"
					If CheckIsMaster=false Then
					  echo "<li class=""backlist"" id=""favtips""><a href=""javascript:void(0)"" onclick=""topicfav(" & id & ",'" & KS.Setting(66) & "'," & BoardID&")"">�ղ�����</a></li>"
					Else 
					   echo "<li class=""backlist"" id=""favtips"" style=""position:relative;"" onMouseOver=""$('#submenu').show();"" onMouseOut=""$('#submenu').hide();""><a href=""#"">��������</a><div id=""submenu"" class=""submenu"">"
					   echo "<dl><a href=""javascript:void(0)"" onclick=""topicfav(" & id & ",'" & KS.Setting(66) & "'," & BoardID&")"">�ղ�����</a></dl>"
					  if verific=1 then
						echo "<dl><a href=""javascript:void(0)"" onclick=""locked("&id &",'" & KS.Setting(66) & "'," & BoardID&")"">��������</a></dl>"
					  else
						echo "<dl><a href=""javascript:void(0)"" onclick=""unlocked("&id &",'" & KS.Setting(66) & "'," & BoardID&")"">�������</a></dl>"
					  end if
						echo "<dl><a href=""javascript:void(0)"" onclick=""delsubject("&id &",'" & KS.Setting(66) & "'," &boardid&")"">ɾ������</a></dl>"
						echo "<dl><a href=""javascript:void(0)"" onclick=""movetopic('" & KS.Setting(66) & "'," & id & ",'" & KS.LoseHtml(subject) & "')"">�ƶ�����</a></dl>"
					  if istop<>0 then
						echo "<dl><a href='javascript:void(0)' onclick=""canceltop(" & ID & ",'" & KS.Setting(66) & "',"&boardid &");"">ȡ���ö�</a></dl>"
					  else
						echo "<dl><a href='javascript:void(0)' onclick=""settop(" & ID & ",'" & KS.Setting(66) & "',"&boardid &",1);"">��Ϊ�ö�</a></dl>"
						echo "<dl><a href='javascript:void(0)' onclick=""settop(" & ID & ",'" & KS.Setting(66) & "',"&boardid &",2);"">��Ϊ���ö�</a></dl>"
					  end if
					  if isbest=1 then
						echo "<dl><a href='javascript:void(0)' onclick=""cancelbest(" & ID & ",'" & KS.Setting(66) & "',"&boardid &");"">ȡ������</a></dl>"
					  else
						echo "<dl><a href='javascript:void(0)' onclick=""setbest(" & ID & ",'" & KS.Setting(66) & "',"&boardid &");"">��Ϊ����</a></dl>"
					  end if
				  End If 
					echo "</div>" 
				Case "jing"
						  If CurrentPage=1 Then
							 If isbest=1 Then
								echo "<img style='float:right;right:130px;position:absolute' src='"  &KS.GetDomain & KS.Setting(66) & "/images/jh.gif' align='absmiddle' alt=""�������϶�Ϊ����"" title=""�������϶�Ϊ����"">"
							 End If
							 If IsTop<>0 Then
								echo "<img style='float:right;position:absolute' src='"  &KS.GetDomain & KS.Setting(66) & "/images/zd.gif' align='absmiddle' alt=""�������ö���ʾ"" title=""�������ö���ʾ"">"
							 End If
							End If
				case "showpage"
						   If AllowShow=true Then
							If KS.IsNul(KS.S("UserName")) Then
							  echo KS.GetClubPageList(MaxPerPage,CurrentPage,TotalPut,TopicID,GCls.ClubPreContent)
							Else
							  echo KS.ShowPage(TotalPut,MaxPerPage,"",CurrentPage,false,false)
							End If
						   End If
				Case Else
					ParseNode sToken, sValue
			End Select 
			Parse    = iPosBegin
		End Function 
		
		Sub ParseArea(sTokenName, sTemplate)
					Select Case sTokenName
						Case "loop"
						       Application(KS.SiteSN&"LoopTemplate")=sTemplate
							  If IsObject(XML) Then
								 For Each TopicNode In Xml.DocumentElement.SelectNodes("row")
									 If IsObject(UserXML) Then set UN=UserXml.DocumentElement.SelectSingleNode("row[@username='" & lcase(TopicNode.SelectSingleNode("@username").text) & "']")  Else Set UN=Nothing
									  n=n+1
									  ReplayID=TopicNode.SelectSingleNode("@id").text
									  scan sTemplate
									 I=I+1
									 
								 Next
									Set Un=Nothing
							   
							  End If
						case "replay"
							If KSUser.GetUserInfo("LockOnClub")="1" Then Exit Sub
							If KS.Setting(54)<>"3" And LoginTF=false Then Exit Sub
							sTemplate=Replace(Replace(sTemplate,"{#InstallDir#}",KS.Setting(3)),"{#ClubDir#}",KS.Setting(66))
							scan sTemplate
						   
					End Select 
		End Sub 
		Sub ParseNode(sTokenType, sTokenName)
					Select Case lcase(sTokenType)
					    case "item"
						  select case lcase(sTokenName)
						     case "n" echo n
						     case "floor" echo GetFloor(n)
						     case "pubtime" echo KS.GetTimeFormat1(TopicNode.SelectSingleNode("@replaytime").text,true)
							 case "pubip"
							    Select Case KS.ChkClng(KS.Setting(58))
								   case 1 
									If KSUser.GetUserInfo("ClubSpecialPower")="1" Then echo "Post IP��" & TopicNode.SelectSingleNode("@userip").text
								   case 2
									If KSUser.GetUserInfo("ClubSpecialPower")="1" Or KSUser.GetUserInfo("ClubSpecialPower")="2" Or CheckIsMaster=true Then echo "Post IP��" & TopicNode.SelectSingleNode("@userip").text
								   case 3
									 If TopicNode.SelectSingleNode("@showip").text="0" And KSUser.GetUserInfo("ClubSpecialPower")<>1 and CheckIsMaster=false and TopicNode.SelectSingleNode("@username").text<>KS.C("UserName") Then
									 Else
									  echo "Post IP��" & TopicNode.SelectSingleNode("@userip").text
									 End If
								  End Select
							 case "showauthoronly"
							     If KS.S("UserName")="" Then
			                      echo " | <a href='" & KS.Setting(3) & KS.Setting(66) & "/display.asp?id=" & TopicID &"&username=" & TopicNode.SelectSingleNode("@username").text &"'>ֻ��������</a>"
								  Else
								  echo " | <a href='" & KS.GetClubShowUrl(TopicID)&"'>��ʾȫ������</a>"
								  End If
								  Echo " <a href='" &KS.GetDomain & "space/?" & PostUserID &"/club' target='_blank'>�鿴����������</a>"
								  If N=1 And TreplayNum>2 Then
								   Dim GoUrl
								   Echo "<select style='margin-left:80px' onclick=""if (this.value!=''){location.href=this.value;}""><option value=''>---������ת��---</option>"
								   For I=1 To TreplayNum
								    If I>MaxPerPage Then
									 If i Mod MaxPerPage = 0 Then
									 GoUrl=KS.Setting(3) & KS.Setting(66) & "/display.asp?id=" & TopicID&"&Page=" & (I \ MaxPerPage) &"#" &i
									 Else
									 GoUrl=KS.Setting(3) & KS.Setting(66) & "/display.asp?id=" & TopicID&"&Page=" & (I \ MaxPerPage+1) &"#" &i
									 End If
									Else
									 GoUrl=KS.Setting(3) & KS.Setting(66) & "/display.asp?id=" & TopicID&"#" &i
									End If
								    Echo "<option value='" & GoUrl & "'>" & i & "¥</option>"
								   Next
								   Echo "</select>"
								  End If
							 case "username" echo TopicNode.SelectSingleNode("@username").text
							 case "userid" echo TopicNode.SelectSingleNode("@userid").text
							 case "spaceurl" echo KS.GetSpaceUrl(TopicNode.SelectSingleNode("@userid").text)
							 case "onlineico"
							   If UN Is Nothing Then Exit Sub
							   If UN.SelectSingleNode("@isonline").text="1" Then
			                     echo "<img src='" & KS.GetDomain & "user/images/online.gif' title='��ǰ����' align='absmiddle'/>"
			                   Else
			                     echo "<img src='" & KS.GetDomain & "user/images/notonline.gif' title='��ǰ������' align='absmiddle'/>"
			                   End If
							 case "usersignandbottomad"
							    If UN Is Nothing Then 
								Sign=""
								ElseIf TopicNode.SelectSingleNode("@showsign").text="1" Then 
								 Sign=UN.SelectSingleNode("@sign").text
								Else
								 Sign=""
								End If
							      Dim BottomAD:BottomAD=GetAdByRnd(37)
								  If BottomAD<>"" Then
								   If Sign<>"" Then Sign="<div class=""usersign"">" & KS.CheckXss(Sign) &"</div>"
								   Sign=Sign & "<div class=""bottomad"">" & BottomAD &"</div>"
								  End If
								  If Sign<>"" THEN echo "<tr><td class=""topicleft"" style=""border-bottom:none"">&nbsp;</td><td>" & Ubbcode(sign,n) &"</td></tr>"
							 case "quoteandreply"
							  If Not KS.IsNul(KS.C("UserName")) Then
							      If (N=1 And BSetting(46)="1") Or (N>1 And BSetting(47)="1") Then
								  echo "<img src='" &KS.Setting(3) & KS.Setting(66) &"/images/Icon_2.gif' align='absmiddle'> <a onclick=""comments('" & KS.Setting(66) &"'," & topicid & "," & replayid & "," & boardid & "," & n & "," & PostUserID &")"" href='javascript:void(0);'>����</a> | "
								  End If
								 If TopicNode.SelectSingleNode("@verific").text="1" Then echo "<img src='" &KS.Setting(3) & KS.Setting(66) &"/images/repquote.gif' align='absmiddle'> <a href='#reply' onclick=""reply("&n&",'" & TopicNode.SelectSingleNode("@username").text & "','" & TopicNode.SelectSingleNode("@replaytime").text & "')"">����</a> | <img src='" &KS.Setting(3) & KS.Setting(66) &"/images/fastreply.gif' align='absmiddle'> <a href='#reply' >�ظ�</a> | "
							 End If
							 echo "<img src='" & KS.Setting(3) & "images/good.gif'><a href=""javascript:void(0)"" onclick=""support(" & TopicID & ","& ReplayID &",'" & KS.Setting(66) &"')"">֧��(<span style='color:red' id='supportnum" &ReplayID&"'>" & KS.ChkClng(TopicNode.SelectSingleNode("@support").text) & "</span>)</a> | <img src='" & KS.Setting(3) & "images/bad.gif'><a href=""javascript:void(0)"" onclick=""opposition(" & TopicID & ","& ReplayID &",'" & KS.Setting(66) &"')"">����(<span style='color:#999999' id='oppositionnum" & ReplayID & "'>" & KS.ChkClng(TopicNode.SelectSingleNode("@opposition").text) & "</span>)</a>"
							 case "topicmanagemenu"
							   If CheckIsMaster Then
							     echo "<a href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=verify&topicid=" & TopicID & "&replyid=" &ReplayID &"&boardid=" &boardid&"' onclick=""return(confirm('ȷ����˸ûظ���?'));"">���</a> | "
								 
								 If TopicNode.SelectSingleNode("@verific").text="1" Then
							     echo "<a href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=replylock&topicid=" & TopicID & "&replyid=" & ReplayID & "&boardid=" &boardid&"' onclick=""return(confirm('ȷ�����θ���Ϣ��?'));"">����</a> | "
							     Else
							     Echo "<a href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=replyunlock&topicid=" & TopicID & "&replyid=" & ReplayID & "&boardid=" &boardid&"' onclick=""return(confirm('ȷ��ȡ�����θ���Ϣ��?'));"">����</a> | "
							     End If
							     If N=1 Then
							      Echo "<a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & ReplayID & "&topicid=" & TopicID & "&istopic=1'>�༭����</a> | <a href=""javascript:void(0)"" onclick=""delsubject("&TopicID &",'" & KS.Setting(66) & "'," &boardid&")"">ɾ������</a>"
							     Else
							      echo "<a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & ReplayID & "&topicid=" & TopicID & "&istopic=0&page=" & CurrentPage & "'>�༭</a> | <a onclick=""delreply('" & KS.Setting(66) &"'," & topicid & "," & replayid & "," & boardid & ")"" href='javascript:void(0);'>ɾ��</a>"
							     End If
							  
							  ElseIf KS.ChkClng(BSetting(29))=1 And KSUser.UserName= PostUserName Then
								 If N=1 Then
								  echo "<img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/edit.gif"" align=""absmiddle""/><a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & ReplayID & "&topicid=" & TopicID & "&istopic=1'>�༭����</a>"
								  Else
								  echo "<img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/edit.gif"" align=""absmiddle""/><a href='" & KS.Setting(3) & KS.Setting(66) & "/post.asp?action=edit&bid=" & boardid&"&id=" & replayID & "&topicid=" & TopicID & "&istopic=0&page=" & CurrentPage & "'>�༭</a> "
								  End If
							  End If
							  echo " <a href=""#top""><img border=""0"" src=""" & KS.Setting(3) & KS.Setting(66) & "/images/p_up.gif"" alt=""�ص�����""/>����</a> <a href=""#reply""><img border=""0"" src=""" & KS.Setting(3) & KS.Setting(66) & "/images/p_down.gif"" alt=""�ص��ײ�""/>�ײ�</a> "
							 case "showusermanage"
							   If CheckIsMaster And  Not UN  Is Nothing Then
							           echo "<div style=""margin:8px;text-align:center"">"
									  If UN.SelectSingleNode("@lockonclub").text="1" Then
										echo "<a onclick='return(confirm(""ȷ���Ը��û��������������""))' href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=unlockuser&userid=" & UN.SelectSingleNode("@userid").text &"' style=""font-weight:bold"">�������</a>"
									  Else
										echo "<a onclick='return(confirm(""ȷ���������û���""))' href='" &KS.Setting(3) & KS.Setting(66) &"/ajax.asp?action=lockuser&userid=" & UN.SelectSingleNode("@userid").text &"' style=""font-weight:bold"">�������û�</a>"
									  End If
										echo "|<a href=""javascript:void(0)"" onclick=""delusertopic(" & topicid&"," & currentpage & "," & n  &",'"&postusername &"'," & boardid &",'" & KS.Setting(66) & "')""  style=""font-weight:bold"">ɾ����������</a>"
									  echo "</div>"
							  End If
							 case "content"
							   Dim Content,UserIsLock,Sign,RightAD
							   RightAd=GetAdByRnd(36)
							   If Not KS.IsNul(RightAd) Then echo "<span class=""rightAd"">" & RightAd &"</span>"
							   If Not Un Is Nothing Then UserIsLock=KS.ChkClng(UN.SelectSingleNode("@lockonclub").text) Else UserIsLock=0
								If UserIsLock=1 Then
									if CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" then
									 Content="<div class=""nopurview"">���û�����������ȫ������,�������ǰ���/����Ա���Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(TopicNode.SelectSingleNode("@content").text)
									else
									 Content="<div class=""nopurview"">�Բ��𣬸��û�����������ȫ������!</div>"
									end if
								ElseIf TopicNode.SelectSingleNode("@verific").text="2" then
									if CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" then
									 Content="<div class=""nopurview"">����Ϣ������,�������ǰ���/����Ա���Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(TopicNode.SelectSingleNode("@content").text)
									else
									 Content="<div class=""nopurview"">�Բ��𣬸���Ϣ������!</div>"
									end if
								ElseIf TopicNode.SelectSingleNode("@verific").text="0" then
									if CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" then
									 Content="<div class=""nopurview"">����Ϣδ���,�������ǰ���/����Ա���Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(TopicNode.SelectSingleNode("@content").text)
									else
									 Content="<div class=""nopurview"">�Բ��𣬸���Ϣδ���!</div>"
									end if
								ElseIf N=1 Then  '����
									  
									 If ShowScore<=0 or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" Then
									   Content=TopicNode.SelectSingleNode("@content").text
									 Else
										If LoginTF=false Then
										Content="<div class=""nopurview"">�Բ�������û�е�¼�����ȵ�¼������Ҫ����ִﵽ<font color=red>" & ShowScore & "</font>�ֲ��ܲ鿴��</div>"
										ElseIf Cint(KSUser.GetUserInfo("Score"))<Cint(ShowScore) Then
										Content="<div class=""nopurview"">�Բ������Ļ��ֲ��㣡����Ҫ����ִﵽ<font color=red>" & ShowScore & "</font>�ֲ��ܲ鿴,����ǰ���û���Ϊ<font color=green>" & KSUser.GetUserInfo("Score") &"</font>�֣�</div>"
										Else
										Content=TopicNode.SelectSingleNode("@content").text
										End If
									  End If
									  dim rsp,rept,replyContent : Session("TopicMusicReply")=0
									  If Instr(Content,"[replyview]")<>0 Then
									   rept=0 : Session("TopicMusicReply")=1
									   If Cbool(LoginTF)=true Then 
										set rsp=GCls.Execute("select top 1 id from " & PostTable &" where topicid=" & TopicID & " and username='" & KS.C("UserName") & "'")
										if not rsp.eof then rept=1
										if CheckIsMaster=true or KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" or ksuser.username=TopicNode.SelectSingleNode("@username").text then rept=1
									   End If
									   
									   if rept=1 then
										ReplyContent="<div class=""replytips""><font color=""gray"">��������ֻ��<b>�ظ�</b>��ſ������</font><hr color='#f1f1f1' size='1'>" & KS.CutFixContent(content, "[replyview]", "[/replyview]", 0) & "</div>"
									   else
										ReplyContent="<div class=""nopurview""><img src='" & KS.Setting(3) & KS.Setting(66) & "/images/locked.gif' align='absmiddle'/><font color=""red"">��������ֻ��<b>�ظ�</b>��ſ������</font></div>"
									   end if
									   content=replace(content,KS.CutFixContent(content, "[replyview]", "[/replyview]", 1),ReplyContent)
									  End If
									  If KS.ChkClng(PostType)=1 Then  Content=Content & GetVote(TopicID,"")  'ͶƱ
								  ElseIf TopicNode.SelectSingleNode("@verific").text="1" Then
									 Content=TopicNode.SelectSingleNode("@content").text
								  Else
								   if CheckIsMaster=true  then
									 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 30px;line-height:30px;"">����Ϣδ���,�������ǰ������Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(TopicNode.SelectSingleNode("@content").text)
								   ElseIf Not KSUser.GetUserInfo("ClubSpecialPower")="1" or KSUser.GetUserInfo("ClubSpecialPower")="2" Then
									 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 30px;line-height:30px;"">����Ϣδ���,�������ǹ���Ա���Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(TopicNode.SelectSingleNode("@content").text)
									Else
									Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 50px;line-height:50px; "">��վ������˻���,����Ϣδͨ�����!</div>"
								   End If
								 end if
							   Content=KSR.ScanAnnex(Content)
							   Content=replace(replace(Content,"��#","{"),"#��","}")  '���˿�Ѵ��ǩ
							   Content=Ubbcode(KSR.ReplaceEmot(Content),n)
							  Dim TopAD:TopAD=GetAdByRnd(68)  '���Ӷ������
							  If TopAD<>"" Then
							   Content="<div class=""topad"">" & TopAD &"</div><div class=""clubcontent"" id=""content" & n& """>" & bbimg(Content) & "</div>"
							  Else
							   Content="<div class=""clubcontent"" id=""content" & n& """>" & bbimg(Content) & "</div>"
							  End If
							  echo Content
							  
							  echo "<span class=""threadcommnets"" id=""comment_" & replayid&""">"   '����ģ��
							  echo GetComments(CommentXML,Boardid,replayid,KS.ChkClng(BSetting(44)),CheckIsMaster)
							  echo "</span>"
						     case "userinfo" 
							   If UN Is Nothing Then
							  	  echo "<div class=""userface""><img src='../Images/Face/boy.jpg' width='82' height='90'></div>"
								  echo "<div style='height:26px;padding-left:5px;margin-top:10px;text-align:left'>�� �� �飺�ο�</div>"
								   PostUserName="�ο�" : PostUserID=0
							  Else
							  
							   Dim UserFaceSrc:UserFaceSrc=UN.SelectSingleNode("@userface").text
							   PostUserName=UN.SelectSingleNode("@username").text : PostUserID=UN.SelectSingleNode("@userid").text
							   if lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
							    
								'==================������ʾ��ʼ========================
							   echo "<div class=""bui"" id=""user" & n& """ style=""display:none"" onmouseover=""showPopUserInfo(" & n &")"" onmouseout=""hidPopUserInfo(" & n & ")""><div class=""l""><div id='f" & n &"'></div>"
							   echo "<div style='margin-top:5px;padding-left:2px'><img src='" & KS.GetDomain & "images/user/log/106.gif'><a href='javascript:void(0)' onclick=""addF(event,'" & PostUserName & "')"">��Ϊ����</a> <img src='" & KS.GetDomain & "images/user/mail.gif'><a href='javascript:void(0)' onclick=""sendMsg(event,'" & PostUserName & "')"">������Ϣ</a></div></div>"
							   echo "<div class='r'>"
							   echo "<li class=""line""><a href='" & KS.GetSpaceUrl(PostUserID) & "' target='_blank'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/home.gif' width='16' height='16' border='0' align='absmiddle' alt='������ҳ'></a>��ҳ  |" 
								 If UN.SelectSingleNode("@email").text <> "" Then
								echo "  <a href='mailto:" & UN.SelectSingleNode("@email").text & "' target='_blank'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/email.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�:[ " & UN.SelectSingleNode("@email").text &" ]'></a>�ʼ�" & vbcrlf
								 Else
							    echo "  <a href='#'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/email-gray.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�'></a>�ʼ�" & vbcrlf
								End If
								echo "  |" 
								If UN.SelectSingleNode("@qq").text <> "" and UN.SelectSingleNode("@qq").text <> "0" Then
								echo " <a href='#'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/qq.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����:[ " & UN.SelectSingleNode("@qq").text & " ]'></a>QQ����"
								Else
								echo "  <a href='#'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/qq-gray.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����'></a>QQ����" & vbcrlf
								End If	
								
								echo "</li><li><span>�û�:</span>" & PostUserName &"</li><li><span>�Ա�:</span>" & UN.SelectSingleNode("@sex").text &"</li><li><span>����:</span>" & UN.SelectSingleNode("@score").text & "��</li><li><span>����:</span>" & UN.SelectSingleNode("@prestige").text &" </li>"
							    echo "<li><span>����:</span>" & UN.SelectSingleNode("@postnum").text & "</li><li><span>����:</span>" & UN.SelectSingleNode("@besttopicnum").text &"</li>"
							    echo "<li class=""line""><span>��¼����:</span>" & UN.SelectSingleNode("@logintimes").text & " ��</li><li class=""line""><span>ע��ʱ��:</span>" & UN.SelectSingleNode("@regdate").text &"</li>"
							    echo "<li class=""line""><span>����¼:</span>" & UN.SelectSingleNode("@lastlogintime").text & "</li></div></div>"               
								'==================������ʾ����========================
							   
							    echo "<div onmouseover=""popUserInfo(this," & n & ");""><div class=""userface""><a href='" & KS.GetSpaceUrl(PostUserID) & "' target='_blank'><img onload='if (this.width>130){this.width=130;}' onerror='this.src=""../images/face/boy.jpg""' src='" & UserFaceSrc &"' border='0'/></a></div></div>"
							   If UN.SelectSingleNode("@isonline").text="1" Then
							    echo "<div class=""username"">" & PostUserName & " <span style='color:#ff6600'>��ǰ����</span></div>"
							   Else
							    echo"<div class=""username"">" & PostUserName & " <span style='color:#888888'>��ǰ����</span></div>"
							   End If
							   echo "����:" 
							   If Not KS.IsNul(KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"color")) Then
							   echo "<span style='color:" & KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"color") &"'>" & KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"usertitle") & "</span>"
							   Else
							   echo KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"usertitle")
							   End If
							   echo "<div style='margin;5px;height:10px;'><img src='" & KS.GetDomain & KS.Setting(66) & "/images/" & KS.A_G(UN.SelectSingleNode("@clubgradeid").text ,"ico") & "'></div>"
							   echo "��������:" & UN.SelectSingleNode("@postnum").text &"<br/>"
							   echo "�û�����:" & UN.SelectSingleNode("@score").text &" ��<br/>"
							   echo "��¼����:" & UN.SelectSingleNode("@logintimes").text &" ��<br/>"
							   echo "ע��ʱ��:" & FormatDateTime(UN.SelectSingleNode("@regdate").text,2) &"<br/>"
							   echo "����¼:" & FormatDateTime(UN.SelectSingleNode("@lastlogintime").text,2) &"<br/>"
							  End If
						  end Select
						case "replay" 
							 select case lcase(sTokenName)
							 case "showupfiles"
							   If KS.ChkClng(BSetting(36))=1 Then
								   If LoginTF=true Then
										If KS.IsNul(BSetting(17)) Or KS.FoundInArr(BSetting(17),KSUser.GroupID,",") Then
										  echo "<tr><td><iframe id=""upiframe"" name=""upiframe"" src=""../user/BatchUploadForm.asp?ChannelID=9994&Boardid=" & boardid & """ frameborder=""0"" width=""100%"" height=""20"" scrolling=""no""></iframe></td></tr>"
										End If
								   End If
								End If 
							case "username" echo ksuser.username
							case "userface"
								 Dim UserFace
								 KSUser.UserLoginChecked
								 If Not KS.IsNUL(KSUser.GetUserInfo("UserFace")) Then
								  UserFace=KSUser.GetUserInfo("UserFace") : If Left(UserFace,1)<>"/" And Left(lcase(UserFace),4)<>"http" Then UserFace=KS.GetDomain & UserFace
								 Else
								  UserFace=KS.GetDomain & "images/face/boy.jpg"
								 End If 
								 echo UserFace
							end select
								  
					End Select
		End Sub

		
		
		Public Sub Kesion()
		    Startime=timer() 
			If KS.Setting(56)="0" Then KS.Die "��վ�ѹر�" & KS.Setting(61)
			If KS.Setting(59)="1" Then response.Redirect(KS.Setting(3) & KS.Setting(66) & "/guestbook.asp")
			LoginTF=KSUser.UserLoginChecked
			If Not KS.IsNul(KS.Setting(69)) Then
			  Dim QueryStr:QueryStr=Request.QueryString
			  Dim QArr:QArr=Split(Split(QueryStr,".")(0),"-")
			  If Ubound(Qarr)>=1 Then
			   ID=KS.ChkClng(Qarr(1))
			  Else
			   ID=KS.ChkClng(KS.S("ID"))
			  End If
			  If Ubound(QArr)>=2 Then  
			   CurrentPage = KS.ChkClng(Qarr(2))
			  Else
			   CurrentPage = KS.ChkClng(Request("page")) 
			  End If
			Else
		      ID=KS.ChkClng(KS.S("ID"))
			  CurrentPage = KS.ChkClng(Request("page")) 
			End If
			If CurrentPage<=0 Then CurrentPage=1
			If KS.Setting(114)="" Then Response.Write "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!":response.end
				   FileContent = KSR.LoadTemplate(KS.Setting(160))
				   If KS.IsNul(FileContent) Then FileContent = "ģ�岻����!"
				   FCls.RefreshType = "guestdisplay" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   GetClubPopLogin FileContent
				   Call GetSubject()
				   If BoardID<>0  Then 
				    KS.LoadClubBoard()
				    Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					If Node Is Nothing Then
					 KS.Die "�Ƿ�����!"
					End If
					 BSetting=Node.SelectSingleNode("@settings").text
		             master=Node.SelectSingleNode("@master").text
					 'FileContent=RexHtml_IF(FileContent) '�ȹ������õı�ǩ,���ٱ�ǩ����
				   End If
				   
				    BSetting=BSetting & "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$" :BSetting=Split(BSetting,"$")
					CheckIsMaster=check() '�Ƿ����

					If verific=0 and CheckIsMaster=false Then
					   KS.Die "<script>alert('�Բ���,�����ӻ�û����ˣ�');history.back();</script>"
					End If
					If DelTF=1 and CheckIsMaster=false Then
					   KS.Die "<script>alert('�Բ���������ɾ��!');location.href='" & KS.GetClubListUrl(boardid) & "';</script>"
					End If
					
					Dim GroupPurview:GroupPurview= True : If Not KS.IsNul(BSetting(1)) and (LoginTF=false or KS.FoundInArr(Replace(BSetting(1)," ",""),KSUser.GroupID,",")=false) Then GroupPurview=false
					Dim UserPurview:UserPurview=True : If Not KS.IsNul(BSetting(10)) and (LoginTF=false or KS.FoundInArr(BSetting(10),KSUser.UserName,",")=false) Then UserPurview=false
					If KSUser.GetUserInfo("ClubSpecialPower")="1" Then UserPurview=true:GroupPurview=True
					Dim ScorePurview:ScorePurview=KS.ChkClng(BSetting(11))
					Dim MoneyPurview:MoneyPurview=KS.ChkClng(BSetting(12))
					   
					   FileContent=Replace(FileContent,"{$GetInstallDir}",KS.Setting(3))
					   FileContent=Replace(FileContent,"{$GetSiteUrl}",KS.GetDomain)
					   FileContent=Replace(FileContent,"{$GetClubInstallDir}",KS.Setting(66))
					If ((GroupPurview=false and Not KS.IsNul(BSetting(10))) or (UserPurview=false)) and boardid<>0 Then
					    ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error2")
						GuestTitle="��Ȩ����" : AllowShow=false
					ElseIf KS.ChkClng(KSUser.GetUserInfo("Score"))<ScorePurView And ScorePurView>0 Then
					    ListTemplate=Replace(Replace(LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error5"),"{$Tips}","����<span>" &ScorePurView&"</span>��"),"{$CurrTips}","����<span>" & KSUser.GetUserInfo("Score") & "</span>��")
						
						GuestTitle="��Ȩ����": AllowShow=false
					ElseIf KS.ChkClng(KSUser.GetUserInfo("Money"))<MoneyPurview And MoneyPurview>0 Then
					    ListTemplate=Replace(Replace(LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error5"),"{$Tips}","�ʽ�<span>" &formatnumber(MoneyPurview,2,-1,-1)&"</span>Ԫ"),"{$CurrTips}","�ʽ�<span>" & formatnumber(KSUser.GetUserInfo("money"),2,-1,-1) & "</span>Ԫ")
						GuestTitle="��Ȩ����" : AllowShow=false
					ElseIf  BSetting(0)="0" And LoginTF=false Then
					    ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error1")
					    GuestTitle="��Ȩ����" : AllowShow=false
				    Else
					
					  FileContent=RexHtml_IF(FileContent) '�ȹ������õı�ǩ,���ٱ�ǩ����
					  Dim PostBtnStr:PostBtnStr="<span style=""position:relative;z-index:1000"" onmouseover=""$('#postlist').show()"" onmouseout=""$('#postlist').hide()""><a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/button_post.png""></a><div id=""postlist"" class=""submenu noli"">"
					   PostBtnStr=PostBtnStr&"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/new_post.gif"" align=""absmiddle""/> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & boardid & """>��������</a></dl>"
					   PostBtnStr=PostBtnStr &"<dl><img src=""" & KS.Setting(3) & KS.Setting(66) & "/images/vote.gif"" align=""absmiddle""> <a href=""" & KS.Setting(3) & KS.Setting(66) & "/post.asp?bid=" & BoardID&"&posttype=1"">����ͶƱ</a></dl>"
					   PostBtnStr=PostBtnStr &"</div></span>"
	
					   FileContent=Replace(FileContent,"{$PostButtonAction}",PostBtnStr)
					   FileContent=Replace(FileContent,"{$BoardID}",BoardID)
					   FileContent=Replace(FileContent,"{$TopicID}",TopicID)
					   FileContent=Replace(FileContent,"{$PostTable}",PostTable)
					   FileContent=Replace(FileContent,"{$IsTop}",IsTop)
					   FileContent=Replace(FileContent,"{$Page}",currentpage)
					   AllowShow=true
					   FileContent=Replace(FileContent,"{$GuestTitle}","{@topic.subjectnohtml}")
					   FileContent=KSR.KSLabelReplaceAll(FileContent)
                       GetReplayList:If IsObject(Xml) Then Call GetTopicList(XML)
					   SCan FileContent
					   If Session("PopTips")<>"" Then  Response.write "<script>popShowMessage('" & KS.Setting(3) & KS.Setting(66) & "','" &Session("PopTips") & "');</script>": Session("PopTips")=""
					
					   KS.Die ""
					   
					End If
				   FileContent=Replace(FileContent,"{$GuestTitle}",GuestTitle)
                   FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   FileContent=Replace(Replace(FileContent,"��#","{"),"#��","}")  '��ǩ�滻����
				   FileContent=RexHtml_IF(FileContent)
				   FileContent=Replace(FileContent,"{#ExecutTime}","ҳ��ִ��" & FormatNumber((timer()-startime),5,-1,0,-1) & "�� powered by <a href='http://www.kesion.com' target='_blank'>KesionCMS 7.0</a>")
				   KS.Echo  FileContent
		End Sub
		
		Sub GetSubject()
		  Dim Param
		  If Request("Move")<>"" Then
		    If Request("Move")="next" Then Param=" Where BoardID=" & KS.ChkClng(KS.S("BoardID")) & " and ID>" & ID & " Order By ID" Else Param=" Where  BoardID=" & KS.ChkClng(KS.S("BoardID")) & " and ID<" & ID & " Order By ID desc"
		  Else
		    Param=" Where ID=" & ID
		  End If
		  Set RST=Conn.Execute("Select top 1 ID,Verific,IsBest,IsTop,CategoryID,Subject,Hits,PostTable,PostType,ShowScore,TotalReplay,BoardID,DelTF From KS_GuestBook" & Param)
		  If RST.Eof Then
		   RST.Close:Set RST=Nothing
		   If Request("Move")<>"" Then
		    KS.Die("<script>alert('��û�м�¼�ˣ�');history.back();</script>")
		   Else
		    KS.Die("<script>alert('�Ƿ�������');window.close();</script>")
		   End If
		  End If
		  ID       = RST("ID") : TopicID=ID
		  verific  = RST("Verific"):IsBest = Cint(RST("IsBest")):IsTop = Cint(RST("IsTop")) : CategoryID=KS.ChkClng(RST("CategoryID")):DelTF = KS.ChkClng(RST("DelTf"))
		  Subject  = RST("Subject") : Subject  = replace(replace(subject,"{","��#"),"}","#��") '���˿�Ѵ��ǩ
		  GCls.Execute("Update KS_GuestBook Set Hits=Hits+1 Where ID=" & ID)
		  Hits     = rst("Hits"): PostTable = RST("PostTable") : PostType=RST("PostType")
		  ShowScore = KS.ChkClng(RST("ShowScore"))
		  TreplayNum= KS.ChkClng(RST("TotalReplay"))
		  TotalReplay=TreplayNum+1
		  FCls.RefreshFolderID = RST("BoardID")
		  BoardID=FCls.RefreshFolderID
		  RST.Close : Set RST=Nothing
		  If IsTop<>0 Then
		    If Not IsObject(Application(KS.SiteSN &"TopXML")) Then MustReLoadTopTopic
			Application(KS.SiteSN &"TopXML").DocumentElement.SelectSingleNode("row[@id=" & id&"]/@hits").text=hits
		  End If
		End Sub
		
		Sub GetReplayList()	
		 MaxPerPage=KS.ChkClng(BSetting(21)) : If MaxPerPage=0 Then MaxPerPage=10
		 Dim Param:Param=" DelTF=0 and topicid=" & ID
		 If Request.QueryString("UserName")<>"" Then Param=Param & " And UserName='" & KS.R(KS.S("UserName")) & "'"
		 If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ClubDisplays"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@rootid",3)
				Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
				Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
				Cmd.Parameters.Append cmd.CreateParameter("@totalusetable",200,1,20)
				Cmd.Parameters.Append cmd.CreateParameter("@param",200,1,110)
				'Cmd.Parameters.Append cmd.CreateParameter("@totalput",3,2,4)
				Cmd("@rootid")= ID
				Cmd("@pagenow")=CurrentPage
				Cmd("@pagesize")=MaxPerPage
				Cmd("@totalusetable")=PostTable
				If Not KS.IsNUL(Request.QueryString("UserName")) Then
				 Cmd("@param")=" and DelTF=0 and username='"+KS.S("UserName")+"'"
				Else
				 Cmd("@param")=" and DelTF=0"
				End If
				Set Rs=Cmd.Execute
				'rs.close  'ע�⣺��Ҫȡ�ò���ֵ�����ȹرռ�¼������
				'TotalPut= cmd("@totalput")
				 TotalPut=GCls.Execute("Select Count(1) From " & PostTable& " Where " & Param)(0)
				'rs.open
				If Not RS.Eof Then 
				   Set XML=KS.RsToXml(RS,"row","")
				Else
					KS.AlertHintScript "û�м�¼��!"
				End If
				Rs.close()
				Set Rs=Nothing
				Set Cmd =  Nothing
			   Exit Sub
		Else
			 If TotalReplay=0 Then TotalReplay=1
			 SQLStr=KS.GetPageSQL(PostTable,"id",MaxPerPage,CurrentPage,0,Param,"*")
			 Dim RS:Set RS=conn.Execute(SQLStr)
			 IF RS.Eof And RS.Bof Then 
				  RS.Close:Set RS=Nothing: totalput=0: exit sub
			 Else
					TotalPut= GCls.Execute("Select Count(1) From " & PostTable& " Where " & Param)(0)
					Set XML=KS.RsToXml(RS,"row","")
					RS.Close:Set RS=Nothing
			End IF
		End If
		
	End Sub
		
	Sub GetTopicList(Xml)
		     If CurrentPage=1 Then N=0 Else N=MaxPerPage*(CurrentPage-1)
			 For Each Node In Xml.DocumentElement.SelectNodes("row")
			    If UserNames="" Then
				 UserNames="'" & trim(Node.SelectSingleNode("@username").text) & "'"
				ElseIF KS.FoundInArr(UserNames,"'" & Node.SelectSingleNode("@username").text & "'",",")=false Then
				 UserNames=UserNames & ",'" & trim(Node.SelectSingleNode("@username").text) & "'"
				End If
				If Pids="" Then
					Pids=Node.SelectSingleNode("@id").text
				Else
				    Pids=Pids & "," & Node.SelectSingleNode("@id").text
				End If
			 Next
			 If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ClubUserLists"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@num",3)
				Cmd.Parameters.Append cmd.CreateParameter("@UserNames",202,1,8000)
				Cmd.Parameters.Append cmd.CreateParameter("@UserFields",202,1,300)
				Cmd("@num")=MaxPerPage
				Cmd("@UserNames")= UserNames 
				Cmd("@UserFields")=UserFields
				Set Rs=Cmd.Execute
			 Else
				Set RS=GCls.Execute("Select top " & MaxPerPage & " " & UserFields &" From KS_User Where UserName in(" & UserNames & ")")
			 End If
			 If Not RS.Eof Then Set UserXml=KS.RsToXml(RS,"row","")
			 RS.Close:Set RS=Nothing
			
			 If Pids<>"" And KS.ChkClng(BSetting(44))<>0 Then
				Set RS=GCls.Execute("Select * From KS_GuestComment Where pid in(" & pids & ") order by orderid,id desc")
				If Not RS.Eof Then
				 Set CommentXML=KS.RsToXml(rs,"row","")
				End If
				RS.Close :Set RS=Nothing
			 End If
	End Sub
		
	Function GetFloor(n)
			  select case n
			   case 1 GetFloor="¥��"
			   case 2 GetFloor="ɳ��"
			   case 3 GetFloor="����"
			   case 4 GetFloor="���"
			   case 5 GetFloor="��ֽ"
			   case 6 GetFloor="�ذ�"
			   case else
			   GetFloor=n & "¥"
			  end select
	 End function
	 
	 Private Function bbimg(strText)
		Dim s,re
        Set re=new RegExp
        re.IgnoreCase =true
        re.Global=True
		s=strText
		re.Pattern="<img(.[^>]*)([/| ])>"
		s=re.replace(s,"<img$1/>")
		re.Pattern="<img(.[^>]*)/>"
		s=re.replace(s,"<img$1 onload=""if (this.width>620) this.width=620;"" onclick=""window.open(this.src)"" style=""cursor:pointer""/>")
		bbimg=s
	End Function
	
	
	
%>
 <!--#include file="../ks_cls/ubbfunction.asp"-->
<%		
	 function check()
	 	Dim KSLoginCls
		Set KSLoginCls = New LoginCheckCls1
		If KSLoginCls.Check=true Then
		  check=true
		  Exit function
		else
			Dim KSUser:Set KSUser=New UserCls
			LoginTF=KSUser.UserLoginChecked
			If Cbool(LoginTF)=false Then 
			  check=false
			  exit function
			elseif KSUser.GetUserInfo("ClubSpecialPower")="2" Or KSUser.GetUserInfo("ClubSpecialPower")="1" Then
			  check=true
			  exit function
			else
			   check=KS.FoundInArr(master, KSUser.UserName, ",")
			End If
		end if
	 End function	
	 
	 '�����ȡ���,AdType�������  36 �Ҳ���,37 �ײ����
	 Function GetAdByRnd(ByVal AdType)
	      Dim AdStr:AdStr=KS.Setting(AdType)
	      If KS.IsNul(AdStr) Then Exit Function
		  Dim AdArr:AdArr=Split(AdStr,"@")
		  Dim RandNum,N: N=Ubound(AdArr)+1
          Randomize
          RandNum=Int(Rnd()*N)
          GetAdByRnd=AdArr(RandNum)
	End Function
		
					  
End Class
%>
