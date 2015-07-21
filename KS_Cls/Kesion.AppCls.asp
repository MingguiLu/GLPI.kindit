<!--#include file="../Conn.asp"-->
<!--#include file="../plus/md5.asp"-->
<!--#include file="../Plus/Session.asp"-->
<!--#include file="Kesion.Label.CommonCls.asp"-->
<!--#include file="Kesion.StaticCls.asp"-->
<!--#include file="Kesion.ClubCls.asp"-->
<!--#include file="Kesion.SpaceApp.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Class KesionAppCls
        Private KS,KSUser, KSR,Tp
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR = New Refresh
		End Sub
		Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSR=Nothing
		 Set KSUser=Nothing
		End Sub
        
		
		Public Sub HomePage()
			   Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			   If QueryStrings<>"" And Ubound(Split(QueryStrings,"-"))>=1 Then
				 Call StaticCls.Run()
			   ElseIf Not KS.IsNul(Request.QueryString("do")) Then
			       Select Case lcase(KS.S("DO"))
					  case "reg" reg
					  case "vote" vote
				   End Select
			   Else
				  Dim Template,FsoIndex:FsoIndex=KS.Setting(5)
				  IF Split(FsoIndex,".")(1)<>"asp" Then
					  Response.Redirect KS.Setting(5):Exit Sub
				  Else
					  Template = KSR.LoadTemplate(KS.Setting(110))
					  FCls.RefreshType = "INDEX" '设置刷新类型，以便取得当前位置导航等
					  FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
					  Template=KSR.KSLabelReplaceAll(Template)
				 End IF
				 Response.Write Template  
			  End If
			  Set StaticCls=Nothing
		End Sub
		
		'二级域名
		Public Sub Domain(S)
		   Select Case Lcase(S)
		     Case lcase(KS.Setting(69))     '小论坛
				 dim Club
				 if instr(lcase(request.querystring),lcase(GCls.ClubPreContent))<>0 then
				  set Club=new ClubDisplayCls
				 else
				  set Club=new ClubCls
				 end if
				 Club.kesion
				 Set Club=Nothing
			 case lcase(KS.JSetting(41))   '求职首页
					If KS.JSetting(0)="0" Then KS.Die "<script>alert('本频道已关闭!');location.href='index.asp';</script>"
					Tp = KSR.LoadTemplate(KS.JSetting(10))
					FCls.RefreshType = "JOBINDEX" '设置刷新类型，以便取得当前位置导航等
					FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
					Tp=JLCls.ReplaceLabel(Tp)
					Tp=KSR.KSLabelReplaceAll(Tp)
					KS.Echo Tp
			 case lcase(KS.SSetting(15))   '空间首页
					Tp = KSR.LoadTemplate(KS.SSetting(7))
					FCls.RefreshType = "SpaceINDEX" '设置刷新类型，以便取得当前位置导航等
					FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
					If Trim(Tp) = "" Then Tp = "空间首页模板不存在!"
					Tp=KSR.KSLabelReplaceAll(Tp)
					KS.Echo Tp
			 case else         '空间
			    'ks.die s
				Dim SApp:Set SApp=New SpaceApp
				SApp.Domain=s
				SApp.Kesion
				If SApp.FoundSpace=false Then HomePage
				Set SApp=Nothing
		   End Select
		End Sub
		
		'===================会员注册开始=========================
		sub reg
		  GCls.ComeUrl=request.ServerVariables("HTTP_REFERER")
		  IF KS.Setting(21)=0 Then : Response.Redirect "../../plus/error.asp?action=error&message=" & Server.URLEncode("<li>对不起，本站暂停新会员注册!</li>") :  Response.End
		   If KS.Setting(117)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		   Tp = KSR.LoadTemplate(KS.Setting(117))
		   FCls.RefreshType="UserRegStep1"
		   If Trim(Tp) = "" Then Tp = "模板不存在!"
          Dim UserRegMustFill:UserRegMustFill=KS.Setting(33)
		  Dim ShowCheckEmailTF:ShowCheckEmailTF=true
		  Dim ShowVerifyCodeTF:ShowVerifyCodeTF=false
		 
		 IF KS.Setting(28)="1" Then ShowCheckEmailTF=false
		 IF KS.Setting(27)="1" then ShowVerifyCodeTF=true
		 
		 If KS.Setting(33)="0" Then
		 Tp = Replace(Tp, "{$ShowUserType}", "<input type='hidden' id='groupid' value='2'/>")
		 Tp = Replace(Tp, "{$DisplayUserType}", " style='display:none'")
		 Else
		 Tp = Replace(Tp, "{$ShowUserType}", UserGroupList())
		 Tp = Replace(Tp, "{$DisplayUserType}", "")
		 End If
		 
		 If KS.Setting(32)="1" Then 
		 Tp = Replace(Tp, "{$Show_Detail}", " style='display:none'")
		 Tp = Replace(Tp, "{$Show_DetailTF}", 1)
		 Else
		 Tp = Replace(Tp, "{$Show_Detail}", "")
		 Tp = Replace(Tp, "{$Show_DetailTF}", 2)
		 End If
		 
		 If KS.Setting(148)="1" Then
		 Tp = Replace(Tp, "{$DisplayQestion}", "")
		 Else
		 Tp = Replace(Tp, "{$DisplayQestion}", " style=""display:none""")
		 End If

		 If KS.Setting(149)="1" Then
		 Tp = Replace(Tp, "{$DisplayMobile}", "")
		 Else
		 Tp = Replace(Tp, "{$DisplayMobile}", " style=""display:none""")
		 End If
		 If KS.Setting(143)="1" Then
		 Tp = Replace(Tp, "{$DisplayAlliance}", "")
		 Else
		 Tp = Replace(Tp, "{$DisplayAlliance}", " style=""display:none""")
		 End If
		 
		 If Mid(KS.Setting(161),1,1)="1" Then
		 Dim RndReg:rndReg=GetRegRnd()
		 Tp = Replace(Tp, "{$DisplayRegQuestion}", "")
		 Tp = Replace(Tp, "{$RegQuestion}", GetRegQuestion(RndReg))
		 Tp = Replace(Tp, "{$AnswerRnd}", GetRegAnswerRnd(RndReg))
		 Else
		 Tp = Replace(Tp, "{$DisplayRegQuestion}", " style=""display:none""")
		 Tp = Replace(Tp, "{$RegQuestion}", "")
		 Tp = Replace(Tp, "{$AnswerRnd}", "")
		 End If
		 
		 Tp = Replace(Tp, "{$Show_Question}", KS.Setting(148))
		 Tp = Replace(Tp, "{$Show_Mobile}", KS.Setting(149))
		 If Request("u")<>"" Then
		 Tp = Replace(Tp, "{$UserName}", " value=""" & split(Request("u"),"@")(0) & """")
		 Else
		 Tp = Replace(Tp, "{$UserName}", "")
		 End If
		 If KS.S("Uid")<>"" Then
		  Tp = Replace(Tp, "{$AllianceUser}", " value=""" & KS.S("Uid") & """ readonly")
		  Tp = Replace(Tp, "{$Friend}", " value=""" & KS.S("F") & """")
		 Else
		  Tp = Replace(Tp, "{$AllianceUser}", "")
		  Tp = Replace(Tp, "{$Friend}", "")
		 End If

		 Tp = Replace(Tp, "{$GetUserRegLicense}", Replace(KS.Setting(23),chr(10),"<br/>"))
		 Tp = Replace(Tp,"{$Show_UserNameLimitChar}",KS.Setting(29))
		 Tp = Replace(Tp,"{$Show_UserNameMaxChar}",KS.Setting(30))
		 Tp = Replace(Tp, "{$Show_CheckEmail}", IsShow(ShowCheckEmailTF))
		 Tp = Replace(Tp, "{$Show_VerifyCodeTF}", IsShow(ShowVerifyCodeTF))
	
         Tp = KSR.KSLabelReplaceAll(Tp) '替换函数标签
		 Response.Write Tp  
		end sub
		Function GetRegRnd()
		  Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
		  Dim RandNum,N: N=Ubound(QuestionArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetRegRnd=RandNum
		End Function
		Function GetRegQuestion(ByVal RndReg)
		  Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
		  GetRegQuestion=QuestionArr(rndReg)
		End Function
		Function GetRegAnswerRnd(ByVal RndReg)
		  GetRegAnswerRnd=md5(rndReg,16)
		End Function        '会员类型
		Function UserGroupList()
			If  KS.Setting(33)="0" Then UserGroupList="":Exit Function
			 Dim Node,Tips
			 Call KS.LoadUserGroup()
			 If KS.ChkClng(KS.S("GroupID"))<>0 Then
				Set Node=Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectSingleNode("row[@id=" & KS.S("GroupID") & "]")
				If Not Node Is Nothing Then
				If KS.ChkClng(Node.SelectSingleNode("@showonreg").text)=0 Then KS.Die "<script>alert('对不起，该用户组不允许注册!');</script>"
				UserGroupList="<span style='font-weight:bold;color:#ff6600'>" & Node.SelectSingleNode("@groupname").text &"</span><input type='hidden' value='" & KS.S("GroupID") & "' id='GroupID' name='GroupID'><span style='display:none' id='tips_" &Node.SelectSingleNode("@id").text&"'>" &Node.SelectSingleNode("@descript").text &"</span>"
			    End If 
				Set Node=Nothing
			Else
			  For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row[@showonreg=1 && @id!=1]")
			  If UserGroupList="" Then
			  Tips=Node.SelectSingleNode("@descript").text
			  UserGroupList="<label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"" checked>" & Node.SelectSingleNode("@groupname").text  & "</label><span style='display:none' id='tips_" &Node.SelectSingleNode("@id").text&"'>" &Node.SelectSingleNode("@descript").text &"</span>"
			  Else
			  UserGroupList=UserGroupList & " <label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"">" & Node.SelectSingleNode("@groupname").text & "</label><span style='display:none' id='tips_" &Node.SelectSingleNode("@id").text&"'>" &Node.SelectSingleNode("@descript").text &"</span>"
			  End If
			 Next
			End If
		End Function
		
		Function IsShow(Show)
			If Show =true Then
				IsShow = ""
			Else
				IsShow = " Style='display:none'"
			End If
		End Function		
		
		'===================会员注册结束=====================
		
		
		'投票系统
		Private Sub Vote()
		   Dim ID:ID=KS.ChkClng(KS.S("ID"))
		   If Id=0 Then KS.Die "error!"
		   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		   RS.Open "Select Top 1 * From KS_Vote Where id=" & id,CONN,1,1
		   If RS.Eof And RS.Bof Then
		     RS.Close:Set RS=Nothing
			 KS.Die "error!"
		   End If
		   
		   Dim LoopStr,XML,Node,Str,LC,Xstr,TotalVote
		   
		   '投票操作
		   If KS.S("Action")="dovote" Then
		     If RS("Status")="0" Then
			   KS.Die "<script>alert('该投票已关闭!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 Dim LoginTF:LoginTF=KSUser.UserLoginChecked()
			 Dim GroupIds:GroupIds=RS("GroupIDs")
			 If RS("nmtp")="1" and LoginTF=false Then
	            KS.Die "<script>alert('对不起，只会登录会员才能投票!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 If Not KS.IsNul(GroupIDs) And KS.FoundInArr(GroupIDs, KSUser.GroupID, ",")=False Then
			 	KS.Die "<script>alert('对不起，您所在的会员组不允许投票!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 If RS("TimeLimit")="1" Then
			 	if now<RS("TimeBegin") then KS.Die "<script>alert('对不起，该投票于" & RS("TimeBegin") & "开放！');location.href='?do=vote&id=" & id&"';</script>"
		        if now>RS("TimeEnd") then KS.Die "<script>alert('对不起，该投票已在" & RS("TimeBegin") & "停止！');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 
			 
		     Dim VoteOption,ItemArr,I,UserName
			 VoteOption=KS.FilterIds(KS.S("VoteOption"))
			 If KS.IsNul(VoteOption) Then
			   KS.Die "<script>alert('您没有选择投票项目!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 
			 Dim IPNum:IPNum=KS.ChkClng(RS("IpNum"))
			 Dim IPNums:IPNums=RS("IPNums")
			 If IpNums<>0 Then
			 	If Conn.Execute("Select Count(ID) From KS_PhotoVote Where UserIp='" & KS.GetIP & "' and ChannelID=-1 And InfoID='" & ID & "'")(0)>=IPNums  Then
	             KS.Die "<script>alert('对不起，最多只能投" & IPNums & "次!');location.href='?do=vote&id=" & id&"';</script>"
	             End If
			 End If
			 If IpNum<>0 Then
			 	If Conn.Execute("Select Count(ID) From KS_PhotoVote Where Year(VoteTime)=" & Year(Now) & " and Month(VoteTime)=" & Month(Now) & " and Day(VoteTime)=" & Day(Now) & " and UserIp='" & KS.GetIP & "' and ChannelID=-1 And InfoID='" & ID & "'")(0)>=IPNum  Then
	             KS.Die "<script>alert('对不起，一天最多只能投" & IPNum & "次!');location.href='?do=vote&id=" & id&"';</script>"
	             End If
			 End If
			 
			 ItemArr=Split(VoteOption,",")
		     set XML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 XML.async = false
			 XML.setProperty "ServerHTTPRequest", true 
			 XML.load(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" &id&".xml"))
			 For I=0 To Ubound(ItemArr)
				 Set Node=XML.DocumentElement.SelectSingleNode("voteitem[@id=" & KS.ChkClng(ItemArr(i)) & "]")
				 Node.childNodes(1).text=KS.ChkClng(Node.childNodes(1).text)+1
				 XML.Save(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" &id&".xml"))
			 Next
			 Application(KS.SiteSN&"_Configvoteitem/vote_"&ID)=""
			 If LoginTF=False Then UserName="游客" Else UserName=KSUser.UserName
			 Conn.Execute("Insert Into [KS_PhotoVote]([ChannelID],[ClassID],[InfoID],[VoteTime],[UserName],[UserIP]) Values(-1,'-1','" & ID & "'," & SqlNowString & ",'" & UserName & "','" & KS.GetIP() & "')")

		   End If
		   
		   Dim Tp:Tp = KSR.LoadTemplate(RS("TemplateID"))
		   If KS.IsNul(Tp) Then 
		     KS.Die "您绑定的模板没有内容,请检查!"
		   End If
		   LoopStr=KS.CutFixContent(Tp, "[loop]", "[/loop]", 0)
		   If Not IsObject(XML) Then
		   Set XML=LFCls.GetXMLFromFile("voteitem/vote_"&ID)
		   End If
		   For Each Node In Xml.DocumentElement.SelectNodes("voteitem")
		       Xstr=Xstr & "<set label='" & Node.childNodes(0).text &"' value='" &Node.childNodes(1).text &"' />"
			   TotalVote=TotalVote+KS.ChkClng(Node.childNodes(1).text)
		   Next
		   For Each Node In Xml.DocumentElement.SelectNodes("voteitem")
		       LC=LoopStr
			   If RS("VoteType")="Single" Then
			   LC=Replace(LC,"{@ItemType}","<input type='radio' name='VoteOption' value='"& Node.getAttribute("id") &"' />")
			   Else
			   LC=Replace(LC,"{@ItemType}","<input type='checkbox' name='VoteOption' value='"& Node.getAttribute("id") &"' />")
			   End If
			   LC=Replace(LC,"{@ItemID}",Node.getAttribute("id"))
			   LC=Replace(LC,"{@ItemName}",Node.childNodes(0).text)
			   LC=Replace(LC,"{@Num}",Node.childNodes(1).text)
            
			dim perVote,pstr
			if totalVote=0 Then TotalVote=0.00000001
			perVote=round(Node.childNodes(1).text/totalVote,4)
			pstr="<img src='../images/Default/bar.gif' width='" & round(360*perVote) & "' height='15' align='absmiddle' />"
			perVote=perVote*100
			if perVote<1 and perVote<>0 then
				pstr=pstr & "&nbsp;0" & perVote & "%"
			else
				pstr=pstr & "&nbsp;" & perVote & "%"
			end if			   
			   LC=Replace(LC,"{@Percent}",Pstr)

			   Str=Str & LC
		   Next
		   Tp=Replace(Tp,KS.CutFixContent(Tp, "[loop]", "[/loop]", 1),str)
		   Tp=Replace(Tp,"{$VoteName}",rs("title"))
		   Tp=Replace(Tp,"{$VoteID}",id)
		   Tp=Replace(Tp,"{$VoteItemXML}",Xstr)
		   RS.Close:Set RS=Nothing
		   Tp=KSR.KSLabelReplaceAll(Tp)
           KS.Die Tp
		End Sub
		
End Class
%>