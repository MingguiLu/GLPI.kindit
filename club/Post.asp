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
Set KSCls = New GuestPost
KSCls.Kesion()
Set KSCls = Nothing

Class GuestPost
        Private KS, KSR,KSUser,Templates,Node,BSetting,BoardID,Master
		Private GuestNum,GuestCheckTF,LoginTF,CategoryNode
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
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
		Sub EchoLn(str)
		 Templates=Templates & str & VBCrlf
		End Sub
%>
<!--#include file="../KS_Cls/Kesion.IfCls.asp"-->
<!--#include file="../KS_Cls/ClubFunction.asp"-->
<%
	Public Sub Kesion()
			If KS.Setting(56)="0" Then response.write "本站已关闭论坛功能":response.end
			 GuestCheckTF=KS.Setting(52)
			 GuestNum=KS.Setting(54)
		     Dim FileContent,WriteForm,PostType
		          If KS.Setting(114)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				   FileContent = KSR.LoadTemplate(KS.Setting(115))
				   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
				   GetClubPopLogin FileContent
				   FCls.RefreshType = "guestwrite" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   LoginTF=KSUser.UserLoginChecked
				   BoardID=KS.ChkClng(Request("bid"))
				   PostType=KS.ChkClng(KS.S("PostType"))
				  
				   Session("UploadFileIDs")=""  '附件ID
				   WriteForm=LFCls.GetConfigFromXML("clubpost","/posttemplate/label","post")
				   WriteForm=Replace(WriteForm,"{$GuestNum}",GuestNum)
				   WriteForm=Replace(WriteForm,"{$CodeTF}",CodeTF)
				   WriteForm=Replace(WriteForm,"{$EmotList}",EmotList)
				   WriteForm=Replace(WriteForm,"{$BoardID}",BoardID)
				   
				   Session("Rnd")=KS.MakeRandom(20)
				   if mid(KS.Setting(161),3,1)="1" Then
				     Dim Qid:Qid=GetQuestionRnd
					 Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
					 WriteForm=Replace(WriteForm,"{$Question}",QuestionArr(Qid))
					 Session("Qid")=Qid
				   end If
				   KS.LoadClubBoard
				  If BoardID<>0 Then
				      Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & BoardID &"]") 
					  If Node Is Nothing Then KS.Die "非法参数!"
					  BSetting=Node.SelectSingleNode("@settings").text
					  Master=Node.SelectSingleNode("@master").text
					  If Node.SelectSingleNode("@parentid").text="0" Then
					    KS.Die "<script>alert('不能在一级栏目下发帖!');history.back();</script>"
					  End If
				 End If
				   BSetting=Bsetting& "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				   BSetting=Split(BSetting,"$")
				   Dim SubjectStr
				   If BoardID<>0 Then
				   
				       '编辑帖子
				       If KS.S("Action")="edit" Then
					      '检查有没有编辑帖子权限
					      Dim TopicID:TopicID=KS.ChkClng(KS.S("TopicID"))
						  Dim ReplyID:ReplyID=KS.ChkClng(KS.S("id"))
						  Dim IsTopic:IsTopic=KS.ChkClng(KS.S("IsTopic"))
						  Dim PostTable,Subject,CategoryId,Content,PostUserName
						  if TopicID=0 Or ReplyID=0 Then
						    KS.Die "<script>alert('参数出错!');history.back();</script>"
						  End If
					      Dim RS:Set RS=Conn.Execute("Select top 1 PostTable,Subject,CategoryId,PostType From KS_GuestBook Where ID=" & TopicID)
						  If RS.Eof And RS.Bof Then
						    RS.Close : Set RS=Nothing
						    KS.Die "<script>alert('参数出错!');history.back();</script>"
						  End If
						  PostTable=RS("PostTable")
						  Subject=RS("Subject")
						  CategoryId=RS("CategoryId") : PostType=RS("PostType")
						  RS.Close
						  
						  
						  RS.Open "Select top 1 * From " & PostTable  & " Where ID=" & ReplyID,conn,1,1
						  If RS.Eof And RS.Bof Then
						    RS.Close : Set RS=Nothing
						    KS.Die "<script>alert('参数出错!');history.back();</script>"
						  End If
						  Content=RS("Content")
						  Content=Replace(Content,"[br]",chr(10))
                          Content=Replace(Replace(Content,"{","｛#"),"}","#｝")  '转换科汛标签
                          Subject=Replace(Replace(Subject,"{","｛#"),"}","#｝")  '转换科汛标签
						  PostUserName=RS("UserName")
						  RS.Close :Set RS=Nothing
						  
						  '检查编辑权限
						  If CheckIsMater=false Then
						    If KSUser.UserName<>PostUserName Or KS.ChkClng(BSetting(29))=0 Then
							 FileContent=Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error9"))
							End If
						  End If
						  WriteForm=Replace(WriteForm,"{$Content}",(content))
						  
						  SubjectStr="<input type='hidden' name='replyId' value='" & ReplyID &"'/>"
						  SubjectStr=SubjectStr & "<input type='hidden' name='IsTopic' value='" & IsTopic &"'/>"
						  SubjectStr=SubjectStr & "<input type='hidden' name='topicId' value='" & topicID &"'/>"
						  SubjectStr=SubjectStr & "<input type='hidden' name='page' value='" & KS.ChkClng(KS.S("Page")) &"'/>"
						  SubjectStr=SubjectStr & "<input type='hidden' name='action' value='edit'/>"
					   End If
					  
					    If IsTopic=0 And KS.S("Action")="edit" Then
					     SubjectStr=SubjectStr & "<input name=""Subject" & Session("Rnd") & """ ID=""Subject" & Session("Rnd")&""" type=""hidden"" maxlength=""150"" value=""" & Subject & """>&nbsp;<strong>编辑<span style='color:red'>“"  &Subject & "” </span>的回复</strong>"	
						Else
												  
                          If BSetting(23)<>"0" Then
						   Dim CategoryStr
						   KS.LoadClubBoardCategory
						   For Each CategoryNode In Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectNodes("row[@boardid=" &BoardID &"]")
							if trim(CategoryId)=trim(CategoryNode.SelectSingleNode("@categoryid").text) Then
							CategoryStr=CategoryStr & "<option value='" &CategoryNode.SelectSingleNode("@categoryid").text  & "' selected>" & CategoryNode.SelectSingleNode("@categoryname").text &"</option>"
							Else
							CategoryStr=CategoryStr & "<option value='" &CategoryNode.SelectSingleNode("@categoryid").text  & "'>" & CategoryNode.SelectSingleNode("@categoryname").text &"</option>"
							End If
						   Next
						   If Not KS.IsNul(CategoryStr) Then
							  CategoryStr="<input type='hidden' id='SelectCategoryId' value='" &BSetting(24) & "'/><select name=""CategoryId"" id=""CategoryId""><option value='0'>主题分类</option>"  & CategoryStr &"</select>"
						   End If
                         End If
						
					     SubjectStr=SubjectStr & "<input type=""text"" name=""Subject" & Session("Rnd") & """ ID=""Subject" & Session("Rnd")&""" style=""border:1px solid #cccccc;height:23px;line-height:23px"" size=""60"" maxlength=""150"" value=""" & Subject & """> <span style=""color:#FF0000"">*</span>"	
						 
						 If KSUser.GetUserInfo("ClubSpecialPower")="0"  Then
						  SubjectStr=SubjectStr & " <span style='color:#999'>你当前级别不支持标题使用HTML语法</span>"
						 Else
						  SubjectStr=SubjectStr & " <span style='color:#999'>支持标题使用HTML语法</span>"
						 End If
					     SubjectStr=CategoryStr  & " " & SubjectStr
						End If
				   End If		
				        
					  	  If PostType=1 Then
						     SubjectStr=SubjectStr  &LFCls.GetConfigFromXML("clubpost","/posttemplate/label","postvote")
						      Dim VXML,VNode,ItemStr,TypeOption,TimeLimitStr,ShowLimitTime
							 If KS.S("Action")="edit" Then
							   Dim RSV:Set RSV=Conn.Execute("Select top 1 * From KS_Vote Where TopicID=" & TopicID)
							   If Not RSV.Eof Then
							    Set VXML=LFCls.GetXMLFromFile("voteitem/vote_"&rsv("ID"))
								For Each VNode In VXml.DocumentElement.SelectNodes("voteitem")
								 ItemStr=ItemStr & "<tr><td><input type=""hidden"" name=""votenum"" value=""" & VNode.childNodes(1).text &"""/><input type=""text"" name=""voteitem"" value=""" & VNode.childNodes(0).text & """ size=""43"" class=""textbox"" /></td></tr>"
								Next
								If RSv("VoteType")="Single" Then
							    TypeOption="<option value=""Single"" selected>单选</option><option value=""Multi"">多选</option>"
								Else
							    TypeOption="<option value=""Single"">单选</option><option value=""Multi""  selected>多选</option>"
							    End If
								If RSV("TimeLimit")="1" Then
								 TimeLimitStr="<label><input type='radio' name='timelimit' onclick=""jQuery('#time').hide();"" value='0'>不启用</label><label><input type='radio' name='timelimit' onclick=""jQuery('#time').show();"" value='1' checked>启用</label>"
								 ShowLimitTime=""
								Else
								 TimeLimitStr="<label><input type='radio' name='timelimit' onclick=""jQuery('#time').hide();"" value='0' checked>不启用</label><label><input type='radio' name='timelimit' onclick=""jQuery('#time').show();"" value='1'>启用</label>"
								 ShowLimitTime=" style='display:none'"
								End If
								If RSv("Nmtp")="1" Then
								 SubjectStr=Replace(SubjectStr,"{$Nmtp}"," checked")
								Else
								 SubjectStr=Replace(SubjectStr,"{$Nmtp}","")
								End If
								SubjectStr=Replace(SubjectStr,"{$ValidDays}",datediff("d",rsv("TimeBegin"),rsv("TimeEnd")))
							   End If
							   RSV.CLose : Set RSV=Nothing
							 Else
							  ItemStr="<tr><td><input type=""text"" name=""voteitem"" size=""43"" class=""textbox"" /></td></tr>"
							  ItemStr=ItemStr & "<tr><td><input type=""text"" name=""voteitem"" size=""43"" class=""textbox"" /></td></tr>"
							  ItemStr=ItemStr & "<tr><td><input type=""text"" name=""voteitem"" size=""43"" class=""textbox"" /></td></tr>"
							  TypeOption="<option value=""Single"">单选</option><option value=""Multi"">多选</option>"
							  TimeLimitStr="<label><input type='radio' name='timelimit' onclick=""jQuery('#time').hide();"" value='0'>不启用</label><label><input type='radio' name='timelimit' onclick=""jQuery('#time').show();"" value='1' checked>启用</label>"
							  ShowLimitTime=""
							  SubjectStr=Replace(SubjectStr,"{$Nmtp}","")
							  SubjectStr=Replace(SubjectStr,"{$ValidDays}",7)
							 End If
							    SubjectStr=Replace(SubjectStr,"{$VoteTypeOption}",TypeOption)
							    SubjectStr=Replace(SubjectStr,"{$VoteItem}",ItemStr)
							    SubjectStr=Replace(SubjectStr,"{$TimeLimit}",TimeLimitStr)
							    SubjectStr=Replace(SubjectStr,"{$ShowLimitTime}",ShowLimitTime)
							 
						  End If
						  
				   		WriteForm=Replace(WriteForm,"{$PostSubject}",SubjectStr)
				   		WriteForm=Replace(WriteForm,"{$PostType}",KS.ChkClng(PostType))

				   
					   Dim GroupPurview:GroupPurview= True : If Not KS.IsNul(BSetting(2)) and KS.FoundInArr(Replace(BSetting(2)," ",""),KSUser.GroupID,",")=false Then GroupPurview=false
					   Dim UserPurview:UserPurview=True : If Not KS.IsNul(BSetting(10)) and KS.FoundInArr(BSetting(10),KSUser.UserName,",")=false Then UserPurview=false
					   Dim ScorePurview:ScorePurview=KS.ChkClng(BSetting(11))
					   Dim MoneyPurview:MoneyPurview=KS.ChkClng(BSetting(12))
					   Dim LimitPostNum:LimitPostNum=KS.ChkClng(BSetting(13))
				   If KS.Setting(59)<>"1" Then  '论坛模式判断有没有权限
					   If KSUser.GetUserInfo("LockOnClub")="1" Then
						FileContent=Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error8"))
					   ElseIf ((GroupPurview=false) or (UserPurview=false)) Then
						FileContent=Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error4"))
					   ElseIf (ScorePurView>0 and KS.ChkClng(KSUser.GetUserInfo("Score"))<ScorePurView) Or (MoneyPurview>0 and KS.ChkClng(KSUser.GetUserInfo("Money"))<MoneyPurview) Then
						   FileContent=Replace(FileContent,"{$WriteGuestForm}",Replace(Replace(Replace(Replace(LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error5"),"{$Money}",moneyPurview),"{$Score}",ScorePurview),"{$CurrScore}",KSUser.GetUserInfo("Score")),"{$CurrMoney}",FormatNumber(KSUser.GetUserInfo("Money"),2,-1)))
					   ElseIf 	datediff("n",KSUser.GetUserInfo("RegDate"),now)<KS.ChkClng(bsetting(9)) Then
						FileContent=Replace(Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error7")),"{$Minutes}",KS.ChkClng(bsetting(9)))
					   ElseIf LimitPostNum>0 Then
						 Dim PostNum:PostNum=Conn.Execute("Select count(1) From KS_GuestBook Where BoardId=" & BoardID & " and UserName='" & KSUser.UserName &"' And DateDiff(" & DataPart_D & ",AddTime," & SqlNowString & ")<1")(0)
						 If PostNum>=LimitPostNum Then
						   FileContent=Replace(Replace(Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error6")),"{$LimitPostNum}",LimitPostNum),"{$PostNum}",PostNum)
						 End If
					   End If
				  End If 
				   
				   If (KS.Setting(57)="1" and LoginTF=false) or (BSetting(0)="0" And LoginTF=false) Then
					GCls.ComeUrl=GCls.GetUrl()
 				    FileContent=Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error3"))
                   Else
				    If LoginTF=true Then
					 WriteForm=Replace(WriteForm,"{$UserName}",KSUser.UserName)
					 WriteForm=Replace(WriteForm,"{$User_Enabled}"," readonly ")
					 WriteForm=Replace(WriteForm,"{$UserEmain}",KSUser.GetUserInfo("Email"))
					 WriteForm=Replace(WriteForm,"{$UserHomePage}",KSUser.GetUserInfo("HomePage"))
					 WriteForm=Replace(WriteForm,"{$UserQQ}",KSUser.GetUserInfo("QQ"))
					Else
					 WriteForm=Replace(WriteForm,"{$UserName}","")
					 WriteForm=Replace(WriteForm,"{$User_Enabled}","")
					 WriteForm=Replace(WriteForm,"{$UserEmain}","")
					 WriteForm=Replace(WriteForm,"{$UserHomePage}","http://")
					 WriteForm=Replace(WriteForm,"{$UserQQ}","")
					End If
 				    FileContent=Replace(FileContent,"{$WriteGuestForm}",WriteForm)
 				    FileContent=Replace(FileContent,"{$RndID}",Session("Rnd"))
 				    FileContent=Replace(FileContent,"{$CheckCode}",CheckCode)
				   End If
				   If Request("action")="edit" then
 				    FileContent=Replace(FileContent,"{$GuestTitle}","编辑帖子")
				   else
 				    FileContent=Replace(FileContent,"{$GuestTitle}","发表新主题")
				   end if
				   If KS.ChkClng(BSetting(36))=1 Then
					   If LoginTF=true Then
							If KS.IsNul(BSetting(17)) Or KS.FoundInArr(BSetting(17),KSUser.GroupID,",") Then
							  Dim UpTips:UpTips="<br/>允许上传附件类型：" & BSetting(37) & "<br/>附件大小不超过"& BSetting(38) &" KB"
							  If KS.ChkClng(BSetting(39))<>0 Then UpTips=UpTips & "<br/>本版面限制每天每人上传" &BSetting(39) & "个文件"
							  FileContent=Replace(FileContent,"{$ShowUpFilesTips}", Uptips)
							  FileContent=Replace(FileContent,"{$ShowUpFiles}", "<iframe id=""upiframe"" name=""upiframe"" src=""../user/BatchUploadForm.asp?ChannelID=9994&Boardid=" & boardid & """ frameborder=""0"" width=""100%"" height=""20"" scrolling=""no"" src=""about:blank""></iframe>")
							End If
					   End If
				   End If
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   FileContent=Replace(Replace(FileContent,"｛#","{"),"#｝","}")  '标签替换回来
				   KS.Echo RexHtml_IF(FileContent)
		End Sub
		
		Function GetQuestionRnd()
		  Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
		  Dim RandNum,N: N=Ubound(QuestionArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetQuestionRnd=RandNum
		End Function
		
		
		Function  CheckCode()
		 IF KS.Setting(53)=1 Then
  	      CheckCode="if (myform.Code" & Session("Rnd") &".value==''){" & vbcrlf
	      CheckCode=CheckCode & "alert('请输入附加码！！');" & vbcrlf
	      CheckCode=CheckCode & "myform.Code" & Session("Rnd") &".focus();" & vbcrlf
  	      CheckCode=CheckCode & "return false;" & vbcrlf
	      CheckCode=CheckCode &  "}" & vbcrlf
	    End IF
		If mid(KS.Setting(161),3,1)="1" Then
  	      CheckCode=CheckCode &"if (myform.Answer" & Session("Rnd") &".value==''){" & vbcrlf
	      CheckCode=CheckCode & "alert('请输入您的回答！！');" & vbcrlf
	      CheckCode=CheckCode & "myform.Answer" & Session("Rnd") &".focus();" & vbcrlf
  	      CheckCode=CheckCode & "return false;" & vbcrlf
	      CheckCode=CheckCode &  "}" & vbcrlf
		End If
	   End Function
					  
	   Function CodeTF()
	     if KS.Setting(53)=0 then CodeTF=" style='display:none'"
	   End Function				  

	   
	   Function EmotList()
	        Dim I
			For I=1 To 24
			 if i<10 then
			  EmotList=EmotList &  "<input type=""radio"" name=""txthead"" value=""0" & I & """"
			 else
			  EmotList=EmotList &  "<input type=""radio"" name=""txthead"" value=""" & I & """"
			 end if
			  IF I=1 Then EmotList=EmotList &  " Checked"
			    if i<10 then
				EmotList=EmotList &  " ><img src=""../editor/ubb/images/smilies/default/0" & I & ".gif"" border=""0"" />"
				else
				EmotList=EmotList &  " ><img src=""../editor/ubb/images/smilies/default/" & I & ".gif"" border=""0""/>"
				end if
			  IF I Mod 12=0 Then EmotList=EmotList &  "<BR>"
			Next
	   End Function
	   
	   
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
	   
End Class
%>
