<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Class SpaceApp
        Public  Domain,FoundSpace,Param
        Private KS,UserName,UserID,KSR,Action,ID,Node,CurrPage,TotalPut,MaxPerPage,PageNum
		Private Template,TemplateSub,SubStr,BlogName,KSBCls
		Private Sub Class_Initialize()
		  MaxPerPage=10 : PageNum=1
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSR=Nothing
		 Call CloseConn()
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then FoundSpace=false : EXIT Sub
		    Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			QueryStrings=KS.UrlDecode(QueryStrings)
			Call Show(QueryStrings)
		End Sub
		
		Sub Show(ByVal QueryStrings)
		Dim QSArr:QSArr=Split(QueryStrings,"/")
		If Ubound(QSArr)>=0 Then
		 UserName=KS.DelSQL(QSArr(0))
		 If KS.ChkClng(UserName)=0 Then
		  Param=" Where UserName='" & UserName & "'"
		 Else
		  Param=" Where UserID=" & KS.ChkClng(UserName)
		 End If
		Else
		  Param=" Where [domain]='" & domain & "'"
		End If

		If Ubound(QSArr)>=1 Then Action=QSArr(1)
		If Ubound(QSArr)>=2 Then ID=KS.ChkClng(QSArr(2))
		If Ubound(QSArr)>=3 Then CurrPage=KS.ChkClng(QSArr(3))
		If CurrPage<=0 Then CurrPage=1
		
		Set KSBCls=New BlogCls
		Dim RS
		If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_ShowSpacesss"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@Param",200,1,220)
				Cmd("@param")=param
				Set Rs=Cmd.Execute
		Else
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Blog" &Param,conn,1,1
		End If
		If RS.Eof And RS.Bof Then
		 rs.close:set rs=nothing : FoundSpace=false
		 Exit Sub
		End If
		FoundSpace=true
		UserName=RS("UserName")
		Session("SpaceUserName")=UserName   '����sql��ǩ����
		UserID=RS("UserID")

		Domain=RS("Domain")
		If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
			If RS("Status")=0 Then
			 rs.close:set rs=nothing
			 KS.Die "<script>alert('�ÿռ�վ����δ���!');window.close();</script>"
			elseif RS("Status")=2 then
			 rs.close:set rs=nothing
			 KS.Die "<script>alert('�ÿռ�վ���ѱ�����Ա����!');window.close();</script>"
			end if
		End If
		If KS.FoundInArr(KS.U_G(Conn.Execute("Select top 1 GroupID From KS_User Where UserName='" & UserName & "'")(0),"powerlist"),"s01",",")=false Then 
		 RS.Close : Set RS=Nothing
		 KS.Die "<script>location.href='../company/show.asp?username=" & username & "';</script>"
		End If
		
		'============================��¼���ʴ���������ÿ�============================================
		conn.execute("update KS_Blog Set Hits=Hits+1 Where UserName='" & UserName & "'")
		If KS.C("UserName")<>"" And KS.C("UserName")<>UserName Then
		   Dim RSV:Set RSV=Server.CreateObject("adodb.recordset")
		   RSV.Open "Select top 1 * From KS_BlogVisitor Where UserName='" & UserName & "' and Visitors='" & KS.R(KS.C("UserName")) & "'",conn,1,3
		   If RSV.Eof And RSV.Bof Then
		     RSV.AddNew
			 RSV("UserName")=UserName
			 RSV("Visitors")=KS.C("UserName")
		   End If
		    RSV("AddDate")=Now
			RSV.Update
		    RSV.Close : Set RSV=Nothing
		 End If
		'============================������¼============================================================
		 
		 Dim Xml:Set XML=KS.RsToXml(rs,"row","")
		 If Not IsObject(xml) Then KS.Die "error xml!"
		 Set Node=XML.DocumentElement.SelectSingleNode("row")
		 Set KSBCls.Node=Node
		 KSBCls.UserName=UserName
		 KSBCls.UserID=UserID
		 KSBCls.Domain=Domain
		 RS.Close : Set RS=Nothing
		 Dim TemplateID:TemplateID=Node.SelectSingleNode("@templateid").text
		 If Action<>"" Then template=Template & KSBCls.GetTemplatePath(TemplateID,"TemplateSub")
		 select case Lcase(action)
		   case "blog"
		      KSBCls.Title="����"
			  BlogList
		   case "fresh"
		      KSBCls.Title="������"
			  FreshList
		   case "log" Call BlogLog
		   case "club"
		      KSBCls.Title="�ҵĻ���"
			   ClubList
		   case "album" 		    
		     KSBCls.Title="���"
			 AlbumList
		   case "showalbum" Call ShowAlbum
		   case "group"
		     KSBCls.Title="Ȧ��"
			 GroupList
		   case "friend"
		     KSBCls.Title="����"
			 FriendList
		   case "xx"
		     KSBCls.Title="�ļ�"
			 Call xxList
		   case "info"
		     KSBCls.Title="����"
			 substr=substr & KSBcls.Location("<strong>��ҳ >> ���˵�</strong>")
			 SubStr=SubStr & KSBCls.UserInfo(Template)
		   case "message"
		     KSBCls.Title="����"
		     Call ShowMessage
		   case "intro"
		     KSBCls.Title="��˾����"
			 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��˾���</strong></div>")
			 Dim Irs:Set Irs=Conn.Execute("Select top 1 Intro From KS_EnterPrise Where UserName='" & UserName & "'")
			 if Not Irs.Eof Then
			 SubStr=SubStr & KS.HtmlCode(Irs(0))
		     Else
		       Irs.Close: Set Irs=Nothing
		       KS.AlertHintScript "�Բ��𣬸��û�������ҵ�û���"
			 End If
			 Irs.Close:Set IrS=Nothing
		   case "news" KSBCls.Title="��˾��̬" : GetNews
		   case "shownews" ShowNews
		   case "product"  ProductList
		   case "showproduct" ShowProduct
		   case "ryzs" KSBCls.Title="����֤��" : GetRyzs
		   case "job" JobList
		   case "showphoto" ShowPhoto
		   case else
		    KSBCls.Title="��ҳ"
		    template=KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
		 end select
		 
		  template=Replace(Template,"{$BlogMain}",replace(SubStr,"{","��#"))
		  template=KSBCls.ReplaceBlogLabel(Template)
		  KS.Echo KSBCls.LoadSpaceHead
		  KS.Echo Replace(Template,"��#","{")
		  
		End Sub
		%>
		<!--#Include file="../ks_cls/ubbfunction.asp"-->
		<%
		'��־
		Sub BlogLog()
		  If ID=0 Then KS.Die "error logid!"
		  Dim RS,i
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
		   RS.Open "Select top 1 * from KS_BlogInfo Where ID=" & ID & " and Status=0",conn,1,1
		  Else
		   RS.Open "Select top 1 * from KS_BlogInfo Where ID=" & ID,conn,1,1
		  End If
		  If RS.EOF And RS.BOF Then
			KS.Die "<script>alert('�������ݳ�������־Ϊ�ݸ壡');history.back();</script>"
		  End If
		  KSBCls.Title=rs("title")
		  
		  If RS("IsTalk")<>"1" Then
		  substr=substr & KSBcls.Location("<strong>��ҳ >> �鿴" & UserName & "����Ĳ���</strong>")
		  Else
		  substr=substr & KSBcls.Location("<strong>��ҳ >> �鿴" & UserName & "�����������</strong>")
		  End If

		  conn.execute("update KS_BlogInfo Set Hits=Hits+1 Where ID=" & ID)
		  iF rs("IsTalk")="1" Then
		      SubStr=SubStr & "<strong><a href='../space/?" & userid & "' target='_blank'>" & UserName & "</a>˵��</strong>" & KS.ReplaceInnerLink(Ubbcode(RS("Content"),i)) & "<BR/><BR/>"
		  Else
			  SubStr=substr & LFCls.GetConfigFromXML("space","/labeltemplate/label","log")
			   Dim EmotSrc:If RS("Face")<>"0" Then EmotSrc="<img src=""../User/images/face/" & RS("Face") & ".gif"" border=""0"">"
			   Dim Tags:Tags=RS("Tags")
			   If Not KS.IsNul(Tags) Then
			   Dim TagList,TagsArr:TagsArr=Split(Tags," ")
					if RS("Tags")<>"" then
					TagList="<div style='display:none'><form id='mytagform' target='_blank' action='../space/?" & username & "/blog' method='post'><input type='text' name='tag' id='tag'></form></div><div style='text-align:left'><strong>��ǩ��</strong>"
					 For I=0 To Ubound(TagsArr)
					  If TagsArr(i)<>"" then
						TagList=TagList &"<a href=""javascript:void(0)"" onclick=""$('#tag').val('" & TagsArr(i) & "');$('#mytagform').submit();"">" & TagsArr(i) & "</a> "
					  end if
					 Next
					 TagList=TagList &"&nbsp;&nbsp;&nbsp;&nbsp;"
					end if
			   End If
				Dim MoreStr:MoreStr="�Ķ�����("&RS("hits")&") | �ظ���("& Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &id)(0) &")</div>"
				Dim ContentStr
				
				Dim JFStr:If RS("Best")="1" then JFStr="  <img src=""../images/jh.gif"" align=""absmiddle"">" else JFStr=""
	
				If KS.IsNul(RS("PassWord")) Then 
				
				 ContentStr=RS("Content")
				ElseIf KS.S("Pass")<>"" Then
				  If KS.S("Pass")=rs("password") then
				   ContentStr=RS("Content")
				  Else
				   SubStr="<br /><br />������,���������־��������!<a href='javascript:history.back(-1)'>����</a><br/>"
				   exit sub
				  End if
				Else
				 SubStr="<br/><br/><br/><form method='post' action='" & KSBCls.GetLogUrl(RS) & "'>��ƪ�����ѱ����˼�����,��������־�Ĳ鿴���룺<input style='border-style: solid; border-width: 1' type='password' name='pass' size='15'>&nbsp;<input type='submit' value=' �鿴 '></form>"
				  exit sub
				End IF
			   SubStr=Replace(SubStr,"{$ShowLogTopic}",EmotSrc & RS("Title") & jfstr)
			   SubStr=Replace(SubStr,"{$ShowLogInfo}","[" & RS("AddDate") & "|by:" & RS("UserName") & "]")
			   SubStr=Replace(SubStr,"{$ShowLogText}",KS.ReplaceInnerLink(Ubbcode(ContentStr,i)))
			   SubStr=Replace(SubStr,"{$ShowLogMore}", TagList&MoreStr)
			   
			   SubStr=Replace(SubStr,"{$ShowTopic}",RS("Title"))
			   SubStr=Replace(SubStr,"{$ShowAuthor}",RS("UserName"))
			   SubStr=Replace(SubStr,"{$ShowAddDate}",RS("AddDate"))
			   SubStr=Replace(SubStr,"{$ShowEmot}",EmotSrc)
			   SubStr=Replace(SubStr,"{$ShowWeather}",KSBCls.GetWeather(RS))
			   SubStr=KSR.ReplaceEmot(SubStr)
			   SubStr=SubStr & "<div style=""padding-left:20px;text-align:left"">��һƪ:" & ReplacePrevNextArticle(ID,"Prev")
			   SubStr=SubStr & "<br>��һƪ:" & ReplacePrevNextArticle(ID,"Next") & "</div><br>"
		   End If
		   Dim Title:Title=RS("Title")

		   RS.Close:Set RS=Nothing
		
	maxperpage=5
	 Dim sqlstr:SqlStr="Select * From KS_BlogComment Where LogID=" & ID & " Order By AddDate Desc"
	 Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open SqlStr,Conn,1,1
     IF Not Rs.Eof Then
		    totalPut = RS.RecordCount
		    If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
				RS.Move (CurrPage - 1) * MaxPerPage
			End If
			Call showContent(rs)
	End If
   If totalput>maxperpage Then
   substr=substr & showpage
   End If
   substr=substr &"<div class=""clear""></div>"
   SubStr=SubStr & "<script src=""writecomment.asp?UserID=" & UserID &"&ID=" & ID & "&UserName=" & UserName & "&Title=" & Title & """></script>"
  
  rs.close:set rs=nothing
End Sub

Sub ShowContent(rs)
     substr=substr & "<div style=""border-bottom:1px solid #f1f1f1;padding-bottom:2px;font-weight:bold;font-size:14px;text-align:left"">&nbsp;&nbsp;������ <font color=red>" & totalPut & " </font> �����ۣ����� <font color=red>" & pagenum & "</font> ҳ,�� <font color=red>" & CurrPage & "</font> ҳ</div>"
    substr=substr & "<table  width='99%' border='0' align='center' cellpadding='0' cellspacing='1'>"
  Dim FaceStr,Publish,i,n
     If CurrPage=1 Then
	 N=TotalPut
	 Else
	 N=totalPut-MaxPerPage*(CurrPage-1)
	 End IF
  Do While Not RS.Eof 
   FaceStr=KS.Setting(3) & "images/face/boy.jpg"

    Publish=RS("AnounName")
	If not Conn.Execute("Select top 1 UserFace From KS_User Where UserName='"& Publish & "'").eof Then
      FaceStr=Conn.Execute("Select top 1 UserFace From KS_User Where UserName='"& Publish & "'")(0)
	  If lcase(left(FaceStr,4))<>"http" and left(FaceStr,1)<>"/" then FaceStr=KS.Setting(3) & FaceStr
   End IF
	
   substr=substr & "<tr>"
   substr=substr & "<td width='70' rowspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;' valign='top'><img width=""50"" height=""52"" src=""" & facestr & """ border=""1"" class=""faceborder"" style=""margin-top:2px;margin-bottom:5px""></td>"
  ' substr=substr & "<td height='25' width=""70%"">"
  ' substr=substr & RS("Title")
  ' substr=substr  & "  </td><td width=""30"" align=""right""><font style='font-size:32px;font-family:""Arial Black"";color:#EEF0EE'> " & N & "</font></td>"
   'substr=substr & "</tr>"
   'substr=substr & "<tr>"
   substr=substr & "<td height='25'>" & ReplaceFace(RS("Content"))
   		 If Not KS.IsNul(RS("Replay")) Then
		 substr=substr&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>����Ϊspace���˵Ļظ�:</b><br>" & RS("Replay") & "<br><div align=right>ʱ��:" & rs("replaydate") &"</div></div>"
         End If
   substr=substr & "	 </td>"
   substr=substr & "</tr>"
   substr=substr & "<tr>"
   
   			 Dim MoreStr,KSUser,LoginTF
				 IF trim(KS.C("UserName"))=trim(RS("UserName")) Then
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>��ҳ</a>| <a href='#'>����</a> | <a href='../User/user_message.asp?Action=CommentDel&id=" & RS("ID") & "' onclick=""return(confirm('ȷ��ɾ����������?'));"">ɾ��</a> | <a href='../user/user_message.asp?id=" & RS("ID") & "&Action=ReplayComment' target='_blank'>�ظ�</a>"
			 Else
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>��ҳ</a>| <a href='#'>����</a> "
			 End If

   substr=substr & "<td align='right' colspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><font color='#999999'>(" & publish & " �����ڣ�" & RS("AddDate") &")</font>&nbsp;&nbsp;" & MoreStr & " </td>"
   substr=substr & "</tr>"
   N=N-1
   RS.MoveNext
		I = I + 1
	  If I >= MaxPerPage Then Exit Do
  loop
 substr=substr & "</table>"

End Sub

Function ReplaceFace(c)
		 Dim str:str="����|Ʋ��|ɫ|����|����|����|����|����|˯|���|����|��ŭ|��Ƥ|����|΢Ц|�ѹ�|��|�ǵ�|ץ��|��|"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K
		 For K=0 To 19
		  c=replace(c,"[e"&K &"]","<img title=""" & strarr(k) & """ src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">")
		 Next
		 ReplaceFace=C
End Function
		
		
		Function ReplacePrevNextArticle(NowID,TypeStr)
		    Dim SqlStr
			If Trim(TypeStr) = "Prev" Then
				   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where istalk<>1 and UserName='" & UserName & "' And ID<" & NowID & " And Status=0 Order By ID Desc"
			ElseIf Trim(TypeStr) = "Next" Then
				   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where istalk<>1 and UserName='" & UserName & "' And ID>" & NowID & " And Status=0 Order By ID Desc"
			Else
				ReplacePrevNextArticle = "":Exit Function
			End If
			 Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
			 RS.Open SqlStr, Conn, 1, 1
			 If RS.EOF And RS.BOF Then
				ReplacePrevNextArticle = "û����"
			 Else
			  ReplacePrevNextArticle = "<a href=""" & KSBCls.GetCurrLogUrl(UserID,RS("ID")) & """ title=""" & RS("Title") & """>" & RS("title") & "</a>"
			 End If
			 RS.Close:Set RS = Nothing
	 End Function
	 
	 '�ҵĻ���
	 Sub ClubList()
	     MaxPerPage=20
	     substr=substr & KSBcls.Location("<strong>��ҳ >> �ҵ���̳����</strong>")
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select id,subject,TotalReplay,BoardID,AddTime,LastReplayTime,LastReplayUser From KS_GuestBook Where deltf=0 and UserName='" & UserName & "' and verific=1 order by ID Desc",conn,1,1
		 If RS.Eof And RS.Bof Then
		  SubStr=SubStr & UserName & "��û�з����κλ��⣡" 
		 Else
				totalPut = RS.RecordCount
				If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				KS.LoadClubBoard
				Dim I:I=0
				SubStr=SubStr & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""border""><tr height=""28"" class=""titlename"">"
				SubStr=SubStr & "<td align=""center"">����</td><td align=""center"">���</td>"
				SubStr=SubStr & "<td align=""center"">�ظ�</td><td align=""center"">��󷢱�</td></tr>"
				Do While Not RS.Eof
				   SubStr=SubStr & "<tr><td class='splittd'><img src='../images/arrow_r.gif' align='absmiddle' /> <a href='" &KS.GetClubShowUrl(rs("id")) & "' target='_blank'>" & replace(replace(rs("subject"),"{","��"),"}","��") & "</a><br/><span class='tips'>����ʱ�䣺" & rs("addTime") & "</span></td>"
				   Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & rs("boardid") &"]")
					if not node is nothing then
						SubStr=SubStr & "<td class='splittd' style='text-align:center'><a href='" & KS.GetClubListUrl(rs("boardid")) &"' target='_blank'>" & Node.SelectSingleNode("@boardname").text & "</a></td>"
					else
						 SubStr=SubStr & "<td class='splittd' style='text-align:center'>---</td>"
					end if
				    Set Node=Nothing
					SubStr=SubStr &"<td style='text-align:center' class='splittd'>" & rs("TotalReplay") & "</td>"
					SubStr=SubStr &"<td style='text-align:center' class='splittd'><a href='../space/?" & RS("LastReplayUser") & "' target='_blank'>" & RS("LastReplayUser") & "</a><br/><span class='tips'>" & rs("LastReplayTime") & "</span></td>"
						   
				   SubStr=SubStr &"</tr>" 
				   I=i+1
				   If I>=MaxPerPage Then Exit Do
				RS.MoveNext
				Loop
				SubStr=SubStr & "</table>"
			End If
		 RS.Close:Set RS=Nothing
		 SubStr=SubStr & vbcrlf & ShowPage
	 End Sub
	 
	 '�������б�
	 Sub FreshList()
	     substr=substr & KSBcls.Location("<strong>��ҳ >> ������</strong>")
		 If KS.C("UserName")=UserName Then
		   Dim RSU:Set RSU=Conn.Execute("Select top 1 userface From KS_User Where UserName='" & UserName & "'")
		   If Not RSU.Eof Then
		       dim myface:myface=rsu(0)
			   If KS.IsNul(myface) Then myface="images/face/boy.jpg"
			   if left(myface,1)<>"/" and lcase(left(myface,4))<>"http" then myface=KS.GetDomain & myface
			   substr=substr &"<div style='text-align:left;color:#888;font-weight:bold'><img src='" & myface & "' align='left' class='faceborder' width='42' height='42' style='margin-right:4px'/>˵˵���ڷ�������..."
			   substr=substr &"<br/><textarea name=""CommentContent"" id=""CommentContent"" cols=""60"" class=""freshtextarea"" onfocus=""freshFocus('CommentContent');"" onblur=""freshBlur('CommentContent');"" rows=""4"">���˵��ʲô���ú�����֪��������顢����ʲô����������10���֣�</textarea><div style=""margin:5px;margin-left:53px;padding-bottom:6px;""><input type='button' value=' �� �� ' onclick=""postsay()"" class='btn'/></div></div>"
		   End If
		   RSU.Close :Set RSU=Nothing
		 End If
		  
		  dim queryParam:queryParam=Request.QueryString
		  dim qparr,ShowType
		  If Not KS.IsNUL(queryParam) Then
		   qparr=split(queryParam,"/")
		   If Ubound(qparr)>=2 Then
		    ShowType=qparr(2)
		   else
		    ShowType=1
		   End If
		  End If
		  If KS.C("UserName")=UserName Then
		  If KS.IsNul(ShowType)  Then ShowType=1
		  substr=substr & "<div class=""mynav""><a href='../space/?" & userid &"/fresh/1/1'"
		  If ShowType="1" Then Substr=substr & " class=""curr"""
		  SubStr=SubStr &">����������</a> <a href='../space/?" & userid &"/fresh/2/1'"
		  If ShowType="2" Then SubStr=subStr &" class=""curr"""
		  SubStr=SubStr &">" & UserName & "��������</a></div>"
		  Else
		   ShowType=2
		  End If
		  MaxPerPage =KSBCls.GetUserBlogParam(UserName,"ContentLen"): If MaxPerPage=0 Then MaxPerPage=10
		  Dim SQLStr
		  If ShowType=1 Then
		  SqlStr="Select f.*,u.userface,u.realname from (KS_BlogInfo f inner join KS_friend ff on f.username=ff.friend) inner join KS_User u on u.username=f.username Where ff.UserName='" & UserName &"' and f.status=0 and f.istalk=1 Order By f.id Desc"
		  Else
		  SqlStr="Select f.*,u.userface,u.realname from KS_BlogInfo f inner join KS_User u on f.username=u.username Where f.UserName='" & UserName &"' and f.status=0 and f.istalk=1 Order By f.id Desc"
		  End If
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		  RSObj.Open SQLStr,Conn,1,1
		   If RSObj.EOF and RSObj.Bof  Then
		     If ShowType=2 Then
		     SubStr=SubStr & "<div>��һ������ʲôҲû��д��</div>"
			 else
		     SubStr=SubStr & "<div>��һ�ĺ��Ѷ�������ʲôҲû��д��</div>"
			 end if
		   Else
		                   totalPut = RSObj.RecordCount
						   If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
							End If
								Dim i:i=0
								substr=substr & "<table width='100%' border='0'>"
								Do While Not RSObj.Eof
								     Dim Uid:Uid=RSObj("UserID")
									 dim userfacesrc:userfacesrc=RSObj("userface")
									 dim rname:rname=rsobj("realname"):if ks.isnul(rname) then rname=rsobj("username")
									 if KS.IsNul(userfacesrc) then userfacesrc="../Images/Face/boy.jpg"
									 if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
								
									 substr=substr & "<tr class=""loglist"">"
									 substr= substr & "<td style='text-align:center;width:54px' class='splittd'><div style='margin:5px 2px 5px 0px'><a href=""../space/?" &uid &""" target=""_blank""><img class='faceborder'  src='" & userfacesrc & "' width='40' height='40' border='0'/></a></div></td><td class='splittd'><a target=""_blank"" href=""../space/?" &rsobj("userid") &""" style=""font-size:16px;color:0F5FBB"">" & rname & "</a> <span style='font-size:14px'>" &  KSR.ReplaceEmot(rsobj("content")) & "</span> " & KS.GetTimeFormat(rsobj("adddate"))& " <a href='" & KS.GetDomain & "space/morefresh.asp' target='_blank'>������</a> "
									 
									 
									 
									 Dim CmtNum:CmtNum=KS.ChkClng(rsobj("totalput"))
									 Dim CmtNumStr:CmtNumStr="(" & CmtNum & ")"
									 If CmtNum>0 Then CmtNumStr="(<span style='color:red'>" & CmtNum & "</span>)"
									 substr=substr & " <a href=""javascript:void(0)"" onclick=""showcmt(" & rsobj("id") & ")""> ����" & CmtNumStr & "</a>"
									 
									 If KS.C("UserName")=UserName And ShowType="2" Then
									   substr=substr & " <a href=""" & KS.GetDomain & "user/user_fresh.asp?action=delfresh&username=" & server.URLEncode(username) &"&id=" & rsobj("id") &""" onclick=""return(confirm('ȷ��ɾ������������'))"" style=""color:#ff6600"">[ɾ��]</a>"
									 End If									 
									 
									 substr=substr & "<div id=""sc" & rsobj("id") & """ style="""
									 If i>0 Then substr=substr & "display:none;"
									 substr=substr &"padding:5px;margin-bottom:6px;margin-left:2px;width:400px;border:1px solid #C1DEFB;background:#E8EFF9;"">"
									 
									 If CmtNum>0 Then
									   Dim RSC:Set RSC=Conn.Execute("Select Top 3 C.AnounName,C.UserName,C.Content,C.Replay,C.replaydate,C.AddDate,U.UserFace,U.UserID,U.RealName From KS_BlogComment C Left Join KS_User U On C.AnounName=u.UserName Where C.LogID=" & KS.ChkClng(rsobj("id")))
									   If Not RSC.Eof Then 
									       Dim UserStr,Urls,Facestr
										   substr=substr & "<table width='100%' cellspacing='0' cellpadding='0'>"
										   substr=substr & "<tr><td class='splittd' colspan='2'>���������¹��� <span style='color:red'>" & CmtNum & "</span> �����ۣ�<a href='../space/?" & uid & "/log/" & KS.ChkClng(rsobj("id")) & "' target='_blank'>�鿴ȫ��...</a></td></tr>"
										   Do While Not RSC.Eof
										    UserStr=RSC("AnounName")
											If KS.IsNul(UserStr) Then UserStr=RSC("UserName")
											UID=KS.ChkClng(RSC("UserID"))
											Dim RealName:RealName=RSC("RealName") : If KS.IsNul(RealName) Then RealName=RSC("AnounName")
											Facestr=RSC("UserFace") : If KS.IsNul(Facestr) Then Facestr="images/face/boy.jpg"
											 if left(Facestr,1)<>"/" and lcase(left(Facestr,4))<>"http" then Facestr="../" & Facestr
											If UID=0 Then Urls="#" Else Urls="../space/?" & UID
										    substr=substr & "<tr><td valign='top' class='splittd' style='width:50px;text-align:center;margin:5px 2px 2px 0px;'><img class='faceborder' src='" &facestr & "' width='40' height='40'/></td><td class='splittd' style=""width:300px""><a href='" & Urls & "'>" & RealName & "</a> " & KS.LoseHtml(RSC("Content")) & " " & KS.GetTimeFormat(RSC("Adddate")) 
											 If Not KS.IsNul(RSC("Replay")) Then
											 substr=substr & "<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>���� <a href='../space/?" & userid & "' target='_blank'>" & rname & "</a> �ظ�:</b><br>" & RSc("Replay") & "<br><div align=right>ʱ��:" & rsc("replaydate") &"</div></div>"
											 End If

											substr=substr & "</td></tr>"
										   RSC.MoveNext
										   Loop
										   substr=substr & "</table>"
									   End If
										   RSC.Close : Set RSC=Nothing
									 End If
									 
									 
									 substr=substr &"<form  name=""form" & rsobj("id") & """ action=""../space/writecomment.asp"" method=""post""><input type=""hidden"" name=""action"" value=""CommentSave""/><input type=""hidden"" name=""id"" value=""" & rsobj("id") & """/><input type=""hidden"" name=""AnounName"" value=""" & KS.C("UserName") & """/><input type=""hidden"" name=""from"" value=""1""/>"
									 substr=substr &"<textarea name=""Content"" id=""c" &rsobj("id") & """ class=""cmttextarea"" onblur=""ThisBlur(" & rsobj("id") & ")"" onfocus=""ThisFocus(" & rsobj("id") & ")"" cols=""50"" rows=""2"">��Ҳ˵һ��...</textarea><br/><div style=""margin:4px 0px 4px 0px""><input type=""submit"" class=""btn"" onclick=""return(postcmt(" & rsobj("id") & "))"" value=""����""/></div></form></div>"
									 
									 
									 
									 
									 

									 substr=substr &"</td></tr>"
								 I = I + 1
								  If I >= MaxPerPage Then Exit Do
								RSObj.MoveNext
								Loop
								substr=substr & "</table>"
			End If
		 RSObj.Close:Set RSObj=Nothing
		 SubStr=SubStr & vbcrlf & ShowPage
	 End Sub
	 
	 '�����б�
	 Sub BlogList()
	     Dim Loc
		 If KS.C("UserName")=UserName Then Loc="<span style=""float:right""><a href=""" & KS.Setting(2) & "/user/User_Blog.asp?Action=Add""><img src='" & KS.Setting(2) & "/user/images/icon7.png' border='0'/>д����</a></span>"
		 Loc=Loc & "<strong>��ҳ >> ����</strong>"
	     substr=substr & KSBcls.Location(loc)
		 MaxPerPage =KSBCls.GetUserBlogParam(UserName,"ListBlogNum"): If MaxPerPage=0 Then MaxPerPage=10
		  
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Param:Param=" IsTalk<>1 and UserName='" & UserName &"'"
		 Dim KeyWord:KeyWord=KS.S("Date") '��������
		 Dim Key:Key=KS.R(KS.S("Key")) '������
		 Dim Tag:Tag=KS.R(KS.S("Tag")) 
		 If IsDate(KeyWord) Then
		       If CInt(DataBaseType) = 1 Then
			   Param=Param & " And i.AddDates>='" & KeyWord & " 00:00:00' and i.AddDatess<='" &KeyWord & " 23:59:59'"
			 else
			   Param=Param & " And i.AddDate>=#" & KeyWord & " 00:00:00# and i.AddDate<=#" &KeyWord & " 23:59:59#"
			 end if
		 End If
		 If ClassID<>0 Then Param=Param & " And i.ClassID=" & ClassID
		 If Key <>"" Then Param=Param & " And i.Title Like '%" & Key & "%'"
		 If Tag <>"" Then Param=Param & " And i.Tags Like '%" & Tag & "%'"
		 
		 If KS.S("Date")<>"" Then substr=substr & "<h2>��������:<font color=red>" & KS.S("Date") & "</font>�Ĳ���</h2></br>"
		 If Tag<>"" Then substr=substr & "<h2>����Tag:<font color=red>" & Tag& "</font>�Ĳ���</h2></br>"
		 iF Key<>"" Then substr=substr & "<h2>�������⺬��""<font color=red>" & Key & "</font>""�Ĳ���</h2></br>"
		 iF ClassID<>0 Then substr=substr & "<h2>�����Զ������ID""<font color=red>" & ClassID & "</font>""�Ĳ���</h2></br>"
		 
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select i.*,t.typename from KS_BlogInfo i inner join ks_blogtype t on i.typeid=t.typeid Where " & Param & " and i.Status=0 Order By i.AddDate Desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				  	  If Key<>"" Then
						substr=substr &"<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>�Ҳ�����־���⺬��<font color=red>""" & key & """</font>�Ĳ���!</p></div>"
						Else
						  If KeyWord="" And ClassID=0 Then
							substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>����û��д���Ĳ��ģ�</p></div>"
						  ElseIf ClassID<>0 Then
							substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>�Ҳ����÷������־!</p></div>"
						  Else
							substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>���ڣ�<font color=red>" & KeyWord & "</font>,��û��д���ģ�</p></div>"
						  End If
					   End if
				 Else
							totalPut = RSObj.RecordCount
							If CurrPage > 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
							End If
							call showlog(RSObj)
			End If
		 RSObj.Close:Set RSObj=Nothing
		 SubStr=SubStr & vbcrlf & ShowPage
	End Sub
    Sub showlog(RS)
		 Dim I,Url,Num
		 Num=(CurrPage-1)*MaxPerPage
		 Do While Not RS.Eof 
		   Url=KSBCls.GetLogUrl(rs)
		   substr=substr & "<div class=""loglist"">"
		   substr= substr & "<a class='title' href='" & Url & "'>" & (Num+i+1) & "��" & rs("title") & "</a> <span class='tips'>" & rs("adddate") & "</span><br/>"
		   substr= substr & "<span class='tips'>���ࣺ[<a target='_blank' href='" & KS.Setting(2) & "/space/morelog.asp?classid=" & rs("typeid") &"'>" & rs("typename") & "</a>] <a href=""" & Url  & """>�Ķ�ȫ��("&RS("hits")&")</a>  <a href=""" & Url  & "#Comment"">�ظ�("& Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &RS("id"))(0) &")</a></span>"
           substr=substr &"</div>"
		  RS.MoveNext
		 I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	End Sub
	
	 '�����б�
	 Sub FriendList()
	     MaxPerPage=20
	     substr=substr & KSBcls.Location("<strong>��ҳ >> " & UserName & "�ĺ���</strong>")
	     substr=substr & "           <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select friend,u.userface,u.username from ks_friend f inner join ks_user u on f.friend=u.username where f.username='" & username & "' and f.accepted=1",Conn,1,1
		                 If RSObj.EOF and RSObj.Bof  Then
						  substr=substr & "<tr><td style=""BORDER: #efefef 1px dotted;text-align:center"" colspan=3>û�мӺ��ѣ�</td></tr>"
						 Else
							   totalPut = RSObj.RecordCount
							   If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrPage - 1) * MaxPerPage
							   End If
								call showfriend(RSObj)
				           End If
		 RSObj.Close:Set RSObj=Nothing
		 substr=substr &  "    </table>" & vbcrlf
		 substr=substr & ShowPage
		End Sub
		
		sub showfriend(RS)
		    Dim I,k
			  Do While Not RS.Eof 
                 substr=substr & "<tr height=""20""> " &vbNewLine
				 for k=1 to 4
				 	   Dim UserFaceSrc:UserFaceSrc=RS("UserFace")
					   if lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
                        substr=substr &"<td width=""25%"" style=""border: #efefef 1px dotted;"" align=""center""><a target=""_blank"" href=""" & KS.GetDomain & "space/?" & RS("username") & """><img width=""80"" height=""80"" src=""" & UserFaceSrc & """ border=""0""></a><div align=""center""><a target=""_blank"" href=""blog.asp?username=" & RS("username") & """ target=""_blank"">" &RS(0) & "</a></div><a href=""javascript:void(0)"" onclick=""ksblog.addF(event,'" & rs("UserName") & "');""><img src=""images/adfriend.gif"" border=""0"" align=""absmiddle"" title=""��Ϊ����"">����</a> <a href=""javascript:void(0)"" onclick=""ksblog.sendMsg(event,'" & rs("username") & "')""><img src=""images/sendmsg.gif"" border=""0"" align=""absmiddle"" title=""��Сֽ��"">��Ϣ</a></td>" & vbnewline
			     RS.MoveNext
			     I = I + 1
				 If I >= MaxPerPage or rs.eof Then Exit for
				 next 
				 do while k<4
				  substr=substr & "<td width=""25%"">&nbsp</td>"
				  k=K+1
				 loop
                 substr=substr & "</tr> " & vbcrlf
				If I >= MaxPerPage Then Exit Do
			 Loop
	end sub
	 
	 'Ȧ���б�
	 Sub GroupList()
	     substr=substr & KSBcls.Location("<strong>��ҳ >> Ȧ��</strong>")
		 MaxPerPage =10
		 substr=substr &"  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select * from KS_team where username='" & username & "' and verific=1 order by id desc",Conn,1,1
		   If RSObj.EOF and RSObj.Bof  Then
			substr=substr & "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3>û�д���Ȧ�ӣ�</td></tr>"
		   Else
				totalPut = RSObj.RecordCount
				If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RSObj.Move (CurrPage - 1) * MaxPerPage
				End If
				ShowGroup(RSObj)
		   End If
		 RSObj.Close:Set RSObj=Nothing
		 substr=substr &  "    </table>" & vbcrlf
		 substr=substr & ShowPage
	 End Sub
			 
	 Sub ShowGroup(RS)		 
		 Dim I
		 Do While Not RS.Eof 
		   substr=substr & "<tr style=""margin:2px;border-bottom:#9999CC dotted 1px;"">"
		   substr=substr & "<td width=""20%"" style=""border-bottom:#9999CC dotted 1px;"">"& vbcrlf
		   substr=substr & " <table style=""BORDER-COLLAPSE: collapse"" borderColor=#c0c0c0 cellSpacing=0 cellPadding=0 border=1>"
		   substr=substr &"	<tr>"
		   substr=substr & "		<td><a href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""><img src=""" & rs("photourl") & """ width=""110"" height=""80"" border=""0""></a></td>"
		   substr=substr & "	 </tr>"
		   substr=substr & " </table>"
		   substr=substr & "</td>"
		   substr=substr & " <td style=""border-bottom:#9999CC dotted 1px;""><a class=""teamname"" href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""> " & rs("TeamName") & "</a><br><font color=""#a7a7a7"">�����ߣ�" & rs("username") & "</font><br><font color=""#a7a7a7"">����ʱ��:" &rs("adddate") & "</font><br>����/�ظ���" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0) & "/" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0) & "&nbsp;&nbsp;&nbsp;��Ա:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "��  </td>"
		   substr=substr & "</tr>"
		   substr=substr & "<tr><td height='2'></td></tr>"
			rs.movenext
			I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
	 
	 Sub AlbumList()
	     SubStr=SubStr & KSBcls.Location("<strong>��ҳ >> ���</strong>")
		 MaxPerPage =9
		 SubStr=SubStr & "  <div class=""albumlist"">" & vbcrlf
		 SubStr=SubStr & "   <ul>" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_Photoxc Where username='" & username & "' and status=1 order by id desc",Conn,1,1
		  If RSObj.EOF and RSObj.Bof  Then
			 substr=substr & "<div style=""border: #efefef 1px dotted;text-align:center"">û�д�����ᣡ</div>"
		  Else
							totalPut = RSObj.RecordCount
							
							If CurrPage>1 And (CurrPage - 1) * MaxPerPage < totalPut Then
								RSObj.Move (CurrPage - 1) * MaxPerPage
							Else
								CurrPage = 1
							End If
							 Dim I,k,Url
							 Do While Not RSObj.Eof 
									  substr=substr & "<li>" &vbcrlf
									          If KS.SSetting(21)="1" Then
											   Url="showalbum-" & RSObj("userid") & "-" & rsobj("id")
											  Else
											   Url="../space/?" & RSObj("userid") &"/showalbum/" &RSObj("id")
											  End If
											  substr=substr &"<div class=""albumbg""><a href=""" & url &""" target=""_blank""><img style=""margin-left:-4px;margin-top:5px"" src=""" &RSObj("photourl") &""" width=""120"" height=""90"" border=0></a></div><B><a href=""" & Url &""">" &RSObj("xcname") &"</a></B> (" & RSObj("xps") & ")<font color=red>[" & GetStatusStr(RSObj("flag")) &"]</font>" & vbcrlf
											  substr=substr &"</li>"
											RSObj.movenext
											I = I + 1
										  If I >= MaxPerPage Then Exit Do
										 Loop
				           End If
		 
		 substr=substr &  "    </ul></div>" & vbcrlf & ShowPage
		 
		 RSObj.Close:Set RSObj=Nothing
	 End Sub
	 

	 
	'�鿴��Ƭ
	 Sub ShowAlbum()
	   	SubStr=SubStr & KSBcls.Location("<strong>��ҳ >> �鿴���</strong>")

	   If ID=0 Then KS.Die "error xcid!"
	    Dim RSXC:Set RSXC=Server.CreateObject("ADODB.RECORDSET")
		RSXC.OPEN "Select top 1 * from ks_photoxc where id=" & id,conn,1,3
		if rsxc.eof and rsxc.bof then
		  rsxc.close:set rsxc=nothing
		  KS.Die "<script>alert('�������ݳ���!');history.back();</script>"
		end if
	   If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
		If RSxc("Status")=0 Then
		 KS.Die "<script>alert('�������δ���!');window.close();</script>"
		elseif RSxc("Status")=2 then
		 KS.Die "<script>alert('������ѱ�����Ա����!');window.close();</script>"
		end if
	   End If
	   KSBCls.Title=rsxc("xcname")
	   Select Case rsxc("flag")
		   Case 1,2
		    If rsxc("Flag")=2 and KS.C("UserName")="" then
			  SubStr=SubStr &"<br><br>��������û�Ա�ɼ�������<a href=""../User/"" target=""_blank"">��¼</a>��"
			Else
			  GetAlbumBody rsxc("xcname")
		    End If
		  Case 3
		    If KS.S("Password")=rsxc("password") or Session("xcpass")=rsxc("password") then
			   Session("xcpass")=KS.S("Password")
			   GetAlbumBody rsxc("xcname")
			else
		      SubStr=SubStr &"<form action=""../space/?" & username &"/showalbum/" & id& """ method=""post"" name=""myform"" id=""myform"">������鿴���룺<input type=""password"" name=""password"" size=""12"" style='border-style: solid; border-width: 1'>&nbsp;<input type='submit' value=' �鿴 '></form>"
		   end if
		  Case 4
		    If KS.C("UserName")=rsxc("username") then
			  GetAlbumBody rsxc("xcname")
			else
			  SubStr=SubStr &"<br><br><li>�������Ϊ��˽��ֻ��������˲���Ȩ�����!</li><li>�������������ˣ�<a href=""../User/""  target=""_blank"">��¼</a>�󼴿ɲ鿴!</li>"
			end if
		 End Select
		 rsxc("hits")=rsxc("hits")+1
		 rsxc.update
		 rsxc.close:set rsxc=nothing
	 End Sub
	 Sub GetAlbumBody(xcname)
	             Dim TotalNum,RS,prevurl,nexturl
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select * from KS_Photozp Where xcid=" & id &" Order By ID Desc",conn,1,1
				 If RS.EOF And RS.BOF Then
				    RS.Close : Set RS=Nothing
					SubStr=SubStr &"<p>�������û����Ƭ��</p>"
				 Else
				        TotalNum=RS.Recordcount
				        If CurrPage>TotalNum Or CurrPage<=0 Then CurrPage=1
				        RS.Move(CurrPage-1)
						Conn.Execute("Update KS_PhotoZP Set Hits=hits+1 Where id=" & rs("id"))
						If KS.SSetting(21)="1" Then
						   prevurl="showalbum-" & userid & "-" & id & "-" & CurrPage-1
						   nexturl="showalbum-" & userid & "-" & id & "-" & CurrPage+1
						Else
						   prevurl="../space/?" & userid & "/showalbum/" & id & "/" & CurrPage-1
						   nexturl="../space/?" & userid & "/showalbum/" & id & "/" & CurrPage+1
						End If
						SubStr=SubStr &"<div style='height:50px;line-height:50px;text-align:center'>���������Ҽ���ҳ��<a style='padding:3px;border:1px solid #cccccc' href='" & prevurl & "'>��һ��</a> ��<font color=red>" & currpage & "</font>/" & TotalNum & "�� <a style='padding:3px;border:1px solid #cccccc' href='" & nexturl& "'>��һ��</a> <a style='padding:3px;border:1px solid #cccccc' href=""" & RS("PhotoUrl") & """ target=""_blank"">�鿴ԭͼ</a></div><div style='padding-bottom:20px;text-align:center'><strong>�������:</strong>" & xcname &" <strong>���:</strong><font color=red>" & rs("hits") & "</font>�� <strong>��С:</strong>" & round(rs("photosize") /1024,2)  & " KB <strong>�ϴ�ʱ��:</strong>" & rs("adddate") & "</div><div style='text-align:center'><img onmouseover=""upNext(this)"" id=""bigimg"" src='" & RS("PhotoUrl") & "' alt=""" & rs("descript") & """ style='border:1px solid #efefef' onload=""if (this.width>450) this.width=450;""/></div><div style='padding-top:20px;text-align:center'>" & rs("descript") & "</div>"
						substr=substr & "<script>" & vbcrlf
						substr=substr &" function upNext(bigimg){"&vbcrlf
						substr=substr &"var lefturl		= '" & prevurl & "';	var righturl	= '" & nexturl & "';"&vbcrlf
						substr=substr &"var imgurl		= righturl;var width	= bigimg.width;	var height	= bigimg.height;"&vbcrlf
						substr=substr &"bigimg.onmousemove=function(){"&vbcrlf
						substr=substr &"if(event.offsetX<width/2){"
						substr=substr &"bigimg.style.cursor	= 'url(../images/default/arr_left.cur),auto';"&vbcrlf
						substr=substr &"imgurl				= lefturl;}else{"&vbcrlf
						substr=substr &"bigimg.style.cursor	= 'url(../images/default/arr_right.cur),auto';"&vbcrlf
						substr=substr &"imgurl				= righturl;}" &vbcrlf
						substr=substr &"}"&vbcrlf
						substr=substr &"bigimg.onmouseup=function(){top.location=imgurl;}}</script>"
						
		   		       RS.Close:Set  RS=Nothing
			    End If
				 SubStr=SubStr & "<script>document.onkeydown=chang_page;function chang_page(event){var e=window.event||event;var eObj=e.srcElement||e.target;var oTname=eObj.tagName.toLowerCase();if(oTname=='input' || oTname=='textarea' || oTname=='form')return;	event = event ? event : (window.event ? window.event : null);if(event.keyCode==37||event.keyCode==33){location.href='" & prevurl &"'}	if (event.keyCode==39 ||event.keyCode==34){location.href='" & nexturl & "'}}</script>"
		End Sub
		Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="����"
			Case 2:GetStatusStr="��Ա"
			Case 3:GetStatusStr="����"
			Case 4:GetStatusStr="��˽"
		   End Select
			GetStatusStr="<font color=red>" & GetStatusStr & "</font>"
		End Function
		
		'����
		Sub ShowMessage()
		 SubStr=SubStr & KSBcls.Location("<div align=""left""><strong>��ҳ >> ���԰�</strong>(<a href=""#write"">ǩд����</a>)</div>")
		 
		 SubStr=substr & "" &  GetWriteMessage() 
		 MaxPerPage =8
		 SubStr=SubStr &"  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_BlogMessage Where UserName='" & UserName & "' and status=1 Order By AddDate Desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				  RSObj.Close :Set RSObj=Nothing
				 SubStr=SubStr &"<tr><td style=""background:#FBFBFB;padding:10px;border: #efefef 1px dotted;text-align:center"">��û���˸���������Ŷ!</td></tr></table>"
				 	 ExiT Sub

				 Else
						   totalPut = Conn.Execute("Select count(1) From KS_BlogMessage Where UserName='" & UserName & "' and status=1")(0)
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
								Else
										CurrPage = 1
								End If
								call showguest(RSObj)
				           End If
		 
		 substr=substr &  "      </table>" & vbcrlf & ShowPage
		 RSObj.Close:Set RSObj=Nothing
		
		
		End Sub
		
		Sub ShowGuest(rs)
		 Dim I,CommentStr,n
		  CommentStr="<br/><div style='border-bottom:1px solid #f1f1f1;padding-bottom:3px;font-weight:bold;font-size:14px'>&nbsp;&nbsp;���� <font color=red>" & totalPut & " </font> ��������Ϣ������ <font color=red>" & pagenum & "</font> ҳ,�� <font color=red>" & CurrPage & "</font> ҳ</div>"
			If CurrPage=1 Then
			 N=TotalPut
			 Else
			 N=totalPut-MaxPerPage*(CurrPage-1)
			 End IF
		  Dim RSU,FaceStr,Publish,MoreStr,Rname
		  Do While Not RS.Eof 
		   FaceStr=KS.Setting(3) & "images/face/boy.jpg"
		
			Publish=KS.R(RS("AnounName"))
			Set RSU=Conn.Execute("Select top 1 UserFace,userid,UserName,RealName From KS_User Where UserName='"& Publish & "'")
			If Not RSU.Eof Then
			  FaceStr=rsu(0) : Rname=rsu(3) : If KS.IsNul(Rname) Then Rname=RSU(2)
			  Publish="<a href='" & KS.GetDomain & "space/?" & RSU(1) & "' target='_blank'>" & Rname & "</a>"
			  If lcase(left(FaceStr,4))<>"http" And Left(FaceStr,1)<>"/" then FaceStr=KS.Setting(3) & FaceStr
			Else
			  Publish="<a href='#'>" & Publish & "</a>"
		    End IF
			RSU.Close

		   
			 CommentStr=CommentStr & "<tr>"
		   CommentStr=CommentStr & "<td valign='top' style='padding-bottom:10px;margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td rowspan='2' width='80' align='center'><img width=""60"" height=""60"" src=""" & facestr & """ border=""1"" /></td><td> <span style='color:#999'>��" & N & "¥ " & publish & " �����ڣ�" & RS("AddDate") &"</span>" 
			 IF KS.C("UserName")=UserName Then
			  MoreStr="<a href='#'>����</a> | <a href='../User/user_message.asp?Action=MessageDel&id=" & RS("ID") & "' onclick=""return(confirm('ȷ��ɾ����������?'))"">ɾ��</a> | <a href='../user/user_message.asp?id=" & RS("ID") & "&Action=ReplayMessage' target='_blank'>�ظ�</a>"
             Else
			  MoreStr="<a href='#'>����</a>"
			 End If
		   CommentStr=CommentStr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & MoreStr & ""
		   
		   CommentStr=CommentStr &"<br/>"
		   If Not KS.IsNUL(RS("Title")) Then
		   CommentStr=CommentStr & RS("Title") & "<br/>"
		   End If
		   CommentStr=CommentStr & Replace(RS("Content"),chr(10),"<br/>")
		   
		    If Not KS.IsNul(RS("Replay")) Then
			 CommentStr=CommentStr&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>����Ϊspace���˵Ļظ�:</b><br>" & RS("Replay") & "<br><div align=right>ʱ��:" & rs("replaydate") &"</div></div>"
			End If
				 
		   CommentStr=CommentStr & "	 </td></tr></table></td>"
		   CommentStr=CommentStr & "</tr>"
		
		   N=N-1
		   RS.MoveNext
				I = I + 1
			  If I >= MaxPerPage Then Exit Do
		  loop
		 'CommentStr=CommentStr & "</table>"
		 SubStr=SubStr & CommentStr
		End Sub
		
		Function GetWriteMessage()
		
		 If KS.SSetting(25)="0" And KS.IsNul(KS.C("UserName")) Then
		  GetWriteMessage="<div style=""margin:20px""><strong>��ܰ��ʾ��</strong>ֻ�л�Ա�ſ�������,����ǻ�Ա����<a href=""javascript:ShowLogin()"">��¼</a>,���ǻ�Ա����<a href=""../?do=reg"" target=""_blank"">ע��</a>��</div>"
		 Else
		 GetWriteMessage = "<iframe src='about:blank' name='hidframe' id='hidframe' width='0' height='0'></iframe><div style=""clear:both""></div><a name=""write""></a><table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 GetWriteMessage = GetWriteMessage & "<form target=""hidframe"" name=""myform"" action=""../plus/ajaxs.asp?action=MessageSave"" method=""post"">"
		 GetWriteMessage = GetWriteMessage & "<input type=""hidden"" value=""" & UserName & """ name=""UserName"">"
		 GetWriteMessage = GetWriteMessage & "<input type=""hidden"" value="""" name=""scontent"">"
		 GetWriteMessage = GetWriteMessage & "<tr><td class=""comment_write_title""><span style='font-weight:bold;font-size:14px'>������������:</span><br/><textarea class=""msgtextarea"" cols='70' rows='4' id=""Content"" onfocus=""if (this.value=='��Ȼ���ˣ���˳�����仰����...') this.value='';"" name=""Content"" onblur=""if (this.value=='') this.value='��Ȼ���ˣ���˳�����仰����...';"">��Ȼ���ˣ���˳�����仰����...</textarea>"
		'GetWriteMessage = GetWriteMessage & "<iframe id=""Editor"" name=""Editor"" src=""../editor/ubb/basic.html?id=Content"" frameBorder=""0"" marginHeight=""0"" marginWidth=""0"" scrolling=""No"" style=""height:150px;width:550px""></iframe></td>"
		
		GetWriteMessage = GetWriteMessage & " <br/>�ǳƣ�"
		GetWriteMessage = GetWriteMessage & "   <input name=""AnounName"""
		If KS.C("UserName")<>"" Then GetWriteMessage = GetWriteMessage & " readonly"
		GetWriteMessage = GetWriteMessage & " maxlength=""100"" type=""text"" value=""" & KS.C("UserName") & """ id=""AnounName"" style=""background:#FBFBFB;color:#999;border:1px solid #ccc;width:120""/>&nbsp;<font color=red>*</font> ��֤�룺<input type=""text"" name=""VerifyCode"" onclick=""this.value='';getCode()"" style=""background:#FBFBFB;color:#999;border:1px solid #ccc;width:50px""><span id='showVerify'>�����������ȡ</span><br/><input type=""submit"" onclick=""return(CheckForm());""  name=""SubmitComment"" value=""OK�ˣ��ύ����"" class=""btn""/>"
		GetWriteMessage = GetWriteMessage & "    </td>"
		GetWriteMessage = GetWriteMessage & "  </tr>"
		GetWriteMessage = GetWriteMessage & "  </form>"
		GetWriteMessage = GetWriteMessage & "</table>"
		End If
		End Function 
		
		Sub GetNews()
		 Dim SQL,i,param
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��˾��̬</strong></div>")
		 Dim RS:Set RS=Conn.Execute("Select classid,classname from ks_userclass where username='" & UserName & "' and typeid=4 order by orderid")
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) tHEN
		     SubStr=SubStr &"<h3><div>������鿴</div></h3><img width='50' src='images/search.png' align='absmiddle'>"
			 if ID=0 Then
			  SubStr=SubStr &"<a href='../space/?" & userid & "/" & Action & "/'><font color=red>ȫ������</font></a>&nbsp;&nbsp;"
			 else
			  SubStr=SubStr &"<a href='../space/?" & userid & "/" & Action & "/'>ȫ������</a>&nbsp;&nbsp;"
			 end if
			 For I=0 To Ubound(SQL,2)
			   if ID=SQL(0,I) then
			   SubStr=SubStr & "<a href='../space/?" & userid & "/" & action & "/" & SQL(0,i) & "'><font color=red>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</font></a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   else
			   SubStr=SubStr & "<a href='../space/?" & userid & "/" & action & "/" & SQL(0,i) & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   end if
			 Next
		 End If
		 if ID=0 Then
		 SubStr=SubStr &"<h3><div>��������</div></h3>"
		 Else
		 SubStr=SubStr &"<h3><div>" & Conn.Execute("Select ClassName From KS_UserClass Where ClassID=" & ID)(0) & "</div></h3>"
		 End If
		 MaxPerPage=10
		 param=" Where UserName='" & UserName & "'"
		 If ID<>0 Then Param=Param & " and classid=" & id
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,AddDate From KS_EnterPriseNews " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			 SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>û�з�����̬����,��<a href='../user/user_EnterPriseNews.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(-1)
				 Dim K,N,Total,url
				 Total=Ubound(SQL,2)+1
				 For I=0 To Total
					If KS.SSetting(21)="1" Then Url="show-news-" & userid & "-" & sql(0,n) & KS.SSetting(22) Else Url="../space/?" & userid & "/shownews/" & sql(0,n)
					SubStr=SubStr &"<tr>"
					SubStr=SubStr & "<td style=""border-bottom: #efefef 1px dotted;height:22""><img src='../images/arrow_r.gif' align='absmiddle'> <a href='" & url & "' target='_blank'>" & SQL(1,N) & "</a>&nbsp;" & sql(2,n)
					SubStr=SubStr & "</td>"
					N=N+1
					If N>=Total Or N>=MaxPerPage Then Exit For
				   SubStr=SubStr &"</tr>"
				 Next
		 End If
		  SubStr=SubStr &"</table>" 
		  SubStr=SubStr & "<div id=""kspage"">" & ShowPage() & "</div>"
		  
		End Sub
		
		'��ʾ��������
		Sub ShowNews()
		 Dim SQL,i,RS,PhotoUrl,url
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��˾��̬ >> �鿴����</strong></div>")
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_EnterPriseNews Where UserName='" & UserName & "' and ID=" & ID  ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('�������ݳ���');window.close();</script>"
		 Else
		   KSBCls.Title=rs("Title")
		   SubStr=SubStr &"<tr><td align='center' style='color:#ff6600;font-weight:bold;font-size:14px'><div style=""font-weight:bold;text-align:center"">" & rs("title") & "</div></td></tr>"
		   SubStr=SubStr & "<tr><td><div style=""text-align:center"">���ߣ�" & UserName & "&nbsp;&nbsp;&nbsp;&nbsp;ʱ��:" & RS("AddDate") & "</div>"
		   SubStr=SubStr & "<hr size=1><div>" & KS.HTMLCode(rs("content")) & "</div></td></tr>"
		   If KS.SSetting(21)="1" Then Url="news-" & username  Else Url="../space/?" & username & "/news"
		   SubStr=SubStr &"<tr><td><div style='text-align:center'><a href='" & Url & "'>[���ع�˾��̬]</a></div></td></tr>"
		 End If
		 SubStr=SubStr &"</table>"   
         RS.Close:Set RS=Nothing
		End Sub
		
		'��Ʒ�б�
		Function ProductList()
		 Dim SQL,i,param,classUrl
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��Ʒչʾ</strong></div>")
		 Dim RS:Set RS=Conn.Execute("Select classid,classname from ks_userclass where username='" & UserName & "' and typeid=3 order by orderid")
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) tHEN
		     SubStr=SubStr &"<h3><div>������鿴</div></h3><img width='50' src='images/search.png' align='absmiddle'>"
			 If KS.SSetting(21)="1" Then classUrl="product-" & username Else classUrl="../space/?" & UserName & "/product"
			 if ID=0 Then
			  SubStr=SubStr & "<a href='" & classUrl & "'><font color=red>ȫ����Ʒ</font></a>&nbsp;&nbsp;"
			 else
			  SubStr=SubStr &"<a href='" & classUrl & "'>ȫ����Ʒ</a>&nbsp;&nbsp;"
			 end if
			 For I=0 To Ubound(SQL,2)
			   If KS.SSetting(21)="1" Then classUrl="product-" & userid & "-" & SQL(0,I) & ks.SSetting(22) Else classUrl="../space/?" & Userid & "/product/" & SQL(0,i)
			   if ID=SQL(0,I) then
			   SubStr=SubStr & "<a href='" & ClassURL & "'><font color=red>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where verific=1 and classid=" & sql(0,i))(0) &")</font></a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   else
			   SubStr=SubStr & "<a href='" & ClassURL & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where verific=1 and classid=" & sql(0,i))(0) &")</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   end if
			 Next
		 End If
		 if ID=0 Then
		 SubStr=SubStr &"<h3><div>���в�Ʒ</div></h3>"
		 Else
		 SubStr=SubStr &"<h3><div>" & Conn.Execute("Select classname from ks_userclass where classid=" &ID)(0) & "</div></h3>"
		 End If
		 MaxPerpage=12
		 param=" Where verific=1 and Inputer='" & UserName & "'"
		 If ID<>0 Then Param=Param & " and classid=" & id
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,PhotoUrl From KS_Product " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>û�з�����Ʒչʾ,��<a href='../user/user_myshop.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage> 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim K,N,Total,PhotoUrl,Url
				 Total=Ubound(SQL,2)+1
				 For I=0 To Total
				   SubStr=SubStr &"<tr>"
				   For K=1 To 4
					PhotoUrl=SQL(2,N)
					If KS.SSetting(21)="1" Then Url="show-product-" &userid & "-" & sql(0,n) & KS.SSetting(22) Else url="../space/?" & Userid & "/showproduct/" & sql(0,n)
					If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nophoto.gif"
					SubStr=SubStr & "<td align='center'>" 
					SubStr=SubStr & "<a href='" & Url & "' target='_blank'><Img border='0' src='" & PhotoUrl & "' alt='" & SQL(1,N) & "' width='130' height='90' /></a><div style='text-align:center'><a href='" & Url & "'>" & KS.Gottopic(SQL(1,N),20) & "</a></div>"
					SubStr=SubStr & "</td>"
					N=N+1
					If N>=Total Or N>=MaxPerPage Then Exit For
				   Next
				   SubStr=SubStr &"</tr>"
				   If N>=Total  Or N>=MaxPerPage Then Exit For
				 Next
		 End If
		 SubStr=SubStr &"</table>" 
		 SubStr=SubStr & ShowPage()  
		End Function
		
		'�鿴��Ʒ����
		Function ShowProduct()
		 Dim SQL,i,RS,PhotoUrl
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��Ʒչʾ >> �鿴��Ʒ����</strong></div>")
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Product Where inputer='" & UserName & "' and ID=" & ID ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('�������ݳ���');window.close();</script>"
		 Else
		   KSBCls.Title=RS("title")
		   photourl=RS("BigPhoto")
		   If PhotoUrl="" Or IsNull(photourl) Then photourl="../images/nophoto.gif"
		   SubStr=SubStr &"<tr><td align='center' style='color:#ff6600;font-weight:bold;font-size:14px'>" & rs("Title") & "</td></tr>"
		   SubStr=SubStr & "<tr><td align='center'><img  style='max-width:600px;width:600px;width:expression(document.body.clientWidth>600?""600px"":""auto"");overflow:hidden;' src='" & photourl &"' border='0'></td></tr>"
		   SubStr=SubStr & "<tr><td><h3><div>��������</div></h3></td></tr>"
		   SubStr=SubStr & "<tr><td>�� �� �̣�" & RS("ProducerName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>��Ʒ���ࣺ" & KS.C_C(RS("tid"),1) & "</td></tr>"
		   SubStr=SubStr & "<tr><td>��Ʒ�ͺţ�" & RS("ProModel") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>Ʒ��/�̱꣺" & RS("TrademarkName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>�� �� �̣�" & RS("ProducerName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>�� �� �ۣ���" & RS("price_market") & " Ԫ</td></tr>"
		   SubStr=SubStr & "<tr><td>�� Ա �ۣ���" & RS("price_member") & " Ԫ</td></tr>"
		   If KS.C_S(5,21)="1" Then
		   SubStr=SubStr & "<tr><td>���߹���<a target='_blank' href=""" & KS.GetItemURL(5,rs("Tid"),rs("ID"),rs("Fname"))   & """><img src='" & KS.GetDomain & "images/ProductBuy.gif' align='absmiddle' border='0'/></a></td></tr>"
		   End If
		   SubStr=SubStr & "<tr><td><h3><div>��ϸ����</div></h3></td></tr>"
		   SubStr=SubStr & "<tr><td>" & bbimg(KS.HtmlCode(RS("proIntro"))) & "</td></tr>"
		 End If
		 SubStr=SubStr &"</table>"   
         RS.Close:Set RS=Nothing
		End Function
		
		Private Function bbimg(strText)
		Dim s,re
        Set re=new RegExp
        re.IgnoreCase =true
        re.Global=True
		s=strText
		re.Pattern="<img(.[^>]*)([/| ])>"
		s=re.replace(s,"<img$1/>")
		re.Pattern="<img(.[^>]*)/>"
		s=re.replace(s,"<img$1  onclick=""window.open(this.src)"" style='max-width:600px;width:600px;width:expression(document.body.clientWidth>600?""600px"":""auto"");overflow:hidden;'/>")
		bbimg=s
	End Function
		
		'��Ƹ
		Sub JobList()
		   SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��ҵ��Ƹ</strong></div>")
		 If KS.C_S(10,21)="0" Then 
		   Dim Jrs:set Jrs=Conn.Execute("Select top 1 Job From ks_Enterprise where username='" & UserName & "'")
		   If Not Jrs.Eof Then
		    SubStr=SubStr & KS.HTMLCode(Jrs(0))
		   Else
		    Jrs.Close: Set Jrs=Nothing
		    KS.AlertHintScript "�Բ��𣬸��û�������ҵ�û���"
		   End If
		   Jrs.Close
		   Set Jrs=Nothing
		   Exit Sub
		 End If
		 
		 SubStr=SubStr &"<h3><div>��Ƹ��Ϣ</div></h3>"
		 MaxPerPage=5
		 Dim Param,rs,sql
		 param=" Where status=1 and UserName='" & UserName & "'"
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,JobTitle,province,city,workexperience,num,salary,refreshtime,status,intro,sex From KS_Job_ZW " & Param &" order by refreshtime desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			 SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>û�з�����Ƹ��Ϣ,��<a href='../User/User_JobCompanyZW.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim I,K,N,Total,PhotoUrl,url
				 Total=Ubound(SQL,2)
				 For I=0 To Total
				     SubStr=SubStr &"<tr><td style='line-height:180%;padding-top:6px;padding-bottom:8px;border-bottom:1px solid #cccccc;'>"
					 SubStr=SubStr & "<font color=#ff6600>��λ���ƣ�" & sql(1,i) & "</font>&nbsp;&nbsp;<a href='" & JLCls.GetZWUrl(SQL(0,I)) & "' target='_blank'>�������</a><br>�����ص㣺" & SQL(2,I) & "&nbsp;" & SQL(3,I) & "&nbsp;&nbsp;��Ƹ������" & SQL(5,I) & " ��<BR>"
					 SubStr=SubStr& "�������ڣ�" & sql(7,i) & "&nbsp;&nbsp;�Ա�Ҫ��" & SQL(10,I) & "<br>��ϸ���ܣ�" & SQL(9,I) & "</td>"
				     SubStr=SubStr &"</tr>"
				 Next
		 End If
		 SubStr=SubStr &"</table>"
		 SubStr=SubStr & ShowPage
		End Sub
		
		'����֤��
		Sub GetRyzs()
		Dim SQL,i,param,RS
		 Substr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ����֤��</strong></div>")
		 SubStr=SubStr &"<h3><div>����֤��</div></h3>"

		 param=" Where status=1 and UserName='" & UserName & "'"
		 SubStr=SubStr & "<table style='margin-bottom:5px' border='0' width='98%' align='center' cellspacing='1' cellpadding='0' bgcolor='#FFFFFF'>"
		 SubStr=SubStr & "<tr bgcolor='#F3F3F3' align='center'><td width='20%' height='20'>֤����Ƭ</td><td width='24%'>֤������</td><td width='21%'>��֤����</td><td width='17%'>��Ч����</td><td width='18%'>��ֹ����</td></tr>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,FZJG,sxrq,jzrq,photourl From KS_EnterPriseZS " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=6><p>û�з�������֤��,��<a href='../user/user_EnterPriseZS.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim K,N,Total,PhotoUrl,url,BeginDateStr,EndDateStr
		 Total=Ubound(SQL,2)
		 For I=0 To Total
		   if i mod 2=0 then
		    SubStr=SubStr &"<tr bgcolor='#ffffff'>"
		   else
		    SubStr=SubStr & "<tr bgcolor='#f6f6f6'>"
		   end if
		    PhotoUrl=SQL(5,i)
			If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nophoto.gif"
			BeginDateStr=SQL(3,I) :	If Not IsDate(BeginDateStr) Then BeginDateStr=Now
			EndDateStr =SQL(4,I) : If Not IsDate(EndDateStr) Then EndDateStr=Now
		    SubStr=SubStr & "<td width='150' style='height:80px;text-align:center;padding-top:6px;padding-bottom:8px;'>" 
			SubStr=SubStr & "<a href='" & PhotoUrl & "' target='_blank'><Img border='0' src='" & PhotoUrl & "' width='85' height='60'></a>"
			SubStr=SubStr & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & sql(1,i) & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & sql(2,i) & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & year(BeginDateStr) & "��" & month(BeginDateStr) & "��</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & year(EndDateStr) & "��" & month(EndDateStr) & "��</td>"
		    SubStr=SubStr &"</tr>"
		 Next
		 End If
		 SubStr=SubStr &"</table>" 
		 SubStr=SubStr & ShowPage  
		End Sub
		
		'��ʾͼƬ
		Function ShowPhoto()
		 Dim SQL,n,RS,PhotoUrlArr,PhotoUrl,t
		 substr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��Ʒչʾ >> �鿴��Ʒ</strong></div>")
		 substr=substr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Photo Where Inputer='" & UserName & "' and ID=" & ID  ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('�������ݳ���');window.close();</script>"
		 Else
		   KSBCls.Title = rs("title")
		   photourlArr=Split(RS("PicUrls"),"|||")
		   n=CurrPage
		   if n<0 then n=0
		   t=Ubound(PhotoUrlArr)
		   If N>=t Then n=0
		   If t=0 Then t=1
		   PhotoUrl=Split(PhotoUrlArr(N),"|")(1)
		   substr=substr & "<tr><td align='center' class='divcenter_work_on'><div class='fpic'><a href='../space/?" & UserName & "/showphoto/" & ID & "/" & n+1 &"'><img  onload=""var myImg = document.getElementById('myImg'); if (myImg.width >580 ) {myImg.width =580 ;};"" id=""myImg"" src='" & photourl &"' title='�鿴��һ��' border='0'></A></div></td></tr>"
		   substr=substr &"<tr><td height='35' align='center'>�����<Script Src='../item/GetHits.asp?Action=Count&m=2&GetFlag=0&ID=" & ID & "'></Script> �ܵ�Ʊ��<Script Src='../item/GetVote.asp?m=2&ID=" & ID & "'></Script> ͶƱ��<a href='../item/Vote.asp?m=2&ID=" & ID & "'>Ͷ��һƱ</a></td></tr>"
           substr=substr & "<tr><td height='35' align='center'>��" & N+1 & "/" & t & "�� <a href='../space/?" & UserName & "/showphoto/" & ID &"/0'><img src='images/picindex.gif' border='0'></a>&nbsp;<a href='../space/?" & UserName & "/showphoto/" & id &"/" & N-1 & "'><img src='images/picpre.gif' border='0'></a>&nbsp;<a href='../space/?" & UserName & "/showphoto/" & id &"/" & N+1 & "'><img src='images/picnext.gif' border='0'></a>&nbsp;<a href='../space/?" & UserName & "/showphoto/" & id &"/" & t-1 & "'><img src='images/picend.gif' border='0'></a></td></tr>"
		   substr=substr & "<tr><td><span class=""writecomment""><Script Language=""Javascript"" Src=""../plus/Comment.asp?Action=Write&ChannelID=2&InfoID=" &id & """></Script></span></td></tr>"
		   substr=substr & "<tr><td>&nbsp;<Img src='images/topic.gif' align='absmiddle'> <strong>��Ʒ���ۣ�</strong><br><span class=""showcomment""><script src=""../ks_inc/Comment.page.js"" language=""javascript""></script><script language=""javascript"" defer>Page(1,2,'" & ID & "','Show','../');</script><div id=""c_" & ID & """></div><div id=""p_" & ID & """ align=""right""></div> </span></td></tr>"
		 End If
		 substr=substr &"</table>"   
		End Function
		
		
		'��Ϣ��
		Sub xxList()
		If KS.IsNUL(Request.ServerVariables("QUERY_STRING")) Then KS.Die "error"
		Dim QueryParam:QueryParam=Request.ServerVariables("QUERY_STRING")&"////"
		Dim Channelid:ChannelID=KS.ChkClng(Split(QueryParam,"/")(2))
		if channelid=0 then channelid=1
		Dim SQL,K,OPStr,RSC:Set RSC=Conn.Execute("Select ChannelID,itemName From KS_Channel Where ChannelStatus=1 and channelid<>6  And ChannelID<>9 And ChannelID<>10 order by channelid")
		SQL=RSc.GetRows(-1)
		RSc.Close:set RSc=Nothing
		For K=0 To Ubound(SQL,2)
		 if channelid=sql(0,k) then
		 OpStr=OpStr & "<option value='../space/?" & userid & "/xx/" & SQL(0,K) & "' selected>" & SQL(1,K) & "</option>"
		 else
		 OpStr=OpStr & "<option value='../space/?" & userid & "/xx/" & SQL(0,K) & "'>" & SQL(1,K) & "</option>"
		 end if
		Next
	    substr= KSBcls.Location("<strong>��ҳ >> ��Ϣ��</strong>")
		Substr=Substr& "<div style='margin:5px'>��Ϣ����<select name='channelid' onchange=""location.href=this.value"">" & opstr & "</select>&nbsp;&nbsp;&nbsp;</div>"
		 MaxPerPage =20
		 substr=substr & "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf

		 Dim Sqlstr
		 Select Case KS.C_S(ChannelID,6) 
		  Case 1
		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 2
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,photourl from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
          Case 3
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 4
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 5
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 7
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 8
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where inputer='" & UserName & "' Order By AddDate Desc"
		 End Select

		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open SqlStr,Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
					substr=substr & "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>�Ҳ�����Ҫ����Ϣ��</p></td></tr>"
				 Else
							totalPut = RSObj.RecordCount
								  If CurrPage>1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
								  End If
						call showxx(RSObj,channelid)
				End If
		 
		 substr=substr &  "            </table>" & vbcrlf
		 substr=substr & showpage
		 RSObj.Close:Set RSObj=Nothing
	End Sub	
	
	Sub showxx(rs,channelid)
		if KS.C_S(ChannelID,6) =2 then       'ͼƬ��ʾ��ͬ
		   substr=substr & GetUserPhoto(RS,MaxPerPage,ChannelID)
		Else
			 Dim K,SQL
			 Do While Not RS.Eof
				substr=substr & "<tr><td style=""border-bottom: #efefef 1px dotted;height:22""><img src=""../images/arrow_r.gif"" align=""absmiddle""> [" & KS.GetClassNP(RS(2)) & "] <a href='" & KS.GetItemUrl(channelid,RS(2),RS(0),RS(5)) & "' target='_blank'>" & RS(1) & "</a>&nbsp;&nbsp;(" & RS(7) & ")</td></tr>"
				K=K+1
				If K>=MaxPerPage Then Exit Do
				RS.MoveNext
			 Loop
		 End if
	End Sub
	'===========9-30========================
			Function GetUserPhoto(RS,totalPut,ChannelID)
		    Dim I,K,Url
			Dim PerLineNum:PerLineNum=4   'ÿ����ʾ��Ʒ��
			  Do While Not RS.Eof 
              GetUserPhoto=GetUserPhoto & "<tr height=""20""> " &vbNewLine
			  
			  For K=1 To PerLineNum
			  If ChannelID=2 Then
			   Url="../space/?" & UserName & "/showphoto/" & RS(0)
			  Else
			   Url=KS.GetItemUrl(channelid,RS(2),RS(0),RS(5))
			  End If
              GetUserPhoto=GetUserPhoto & "  <td style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted;text-align:center""><a href=""" & Url & """ target=""_blank""><img style='border:1px #efefef solid' width=120 height=80 src=""" & rs("photourl") & """ border=""0""></a><br><a href=""" & Url & """ target=""_blank"">" & KS.Gottopic(RS(1),15) & "</a></td>" & vbnewline
             RS.MoveNext
			    I = I + 1
				If rs.eof or I >= totalPut Then Exit For
			  Next
			   For K=K+1 To PerLineNum
            GetUserPhoto=GetUserPhoto & "   <td width=120 style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted;text-align:center"">&nbsp;</td> " & vbcrlf
			   Next
            GetUserPhoto=GetUserPhoto & "   </tr> " & vbcrlf
				If I >= totalPut Then Exit Do
			 Loop

		End Function
		
		'ͨ�÷�ҳ
		Public Function ShowPage()
		         Dim I, PageStr
				 PageStr = ("<div class=""fenye""><table border='0' align='right'><tr><td><div class='showpage' style='height:28px'>")
					if (CurrPage>1) then pageStr=PageStr & "<a href=""../space/?" & userid & "/" &action & "/" & ID & "/" & CurrPage-1 & """ class=""prev"">��һҳ</a>"
				   if (CurrPage<>PageNum) then pageStr=PageStr & "<a href=""../space/?" & userid & "/" &action & "/" & ID & "/" & CurrPage+1 & """ class=""next"">��һҳ</a>"
				   pageStr=pageStr & "<a href=""../space/?" & userid & "/" &action & """ class=""prev"">�� ҳ</a>"
				
				    If (totalPut Mod MaxPerPage) = 0 Then
						pagenum = totalPut \ MaxPerPage
					Else
						pagenum = totalPut \ MaxPerPage + 1
					End If
					Dim startpage,n,j
					 if (CurrPage>=7) then startpage=CurrPage-5
					 if PageNum-CurrPage<5 Then startpage=PageNum-10
					 If startpage<=0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""../space/?" & userid & "/" &action & "/" &id & "/" & J&""">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""../space/?" & userid & "/" &action & "/" &id & "/" & PageNum&""">ĩҳ</a>"
					 pageStr=PageStr & " <span>��" & totalPut & "����¼,��" & PageNum & "ҳ</span></div></td></tr></table>"
				     PageStr = PageStr & "</div>"
			         ShowPage = PageStr
	     End Function
End Class
%>
