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
		    echo "�޸�" & KS.C_S(Channelid,3)
		  Else
		    echo "����" & KS.C_S(Channelid,3)
		  End If
		case "selectclassid"
		   Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) 
		case "status"
		   if action="Edit" Then
		     If RS("Verific")<>1 Then
			  if rs("verific")=2 Then
		       echo "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""ȷ������Ͷ�巢����?""));}' checked>����ݸ�</label>"
			  else
		       echo "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""ȷ������Ͷ�巢����?""));}' value='2'>����ݸ�</label>"
			  end if
			 Else
			  echo "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
		   Else
		    echo "<label style='color:#999999'><input type='checkbox' name='status' value='2'>����ݸ�</label>"
		   End If
		case "content"
		  If Action="Edit" Then
		   if KS.C_S(ChannelID,6)=1 Then if not KS.IsNul(rs("ArticleContent")) then echo Server.HtmlEncode(rs("ArticleContent"))
		  End If
		case else
		   on error resume next
		   Dim II,DV
		   if instr(sTemp,"|select")<>0 then  '����������
			 For II=0 To Ubound(UserDefineFieldArr,2)
			   If Lcase(UserDefineFieldArr(0,Ii))=Lcase(split(sTemp,"|")(0)) Then
			      If Action="Edit" Then DV=RS(Trim(UserDefineFieldArr(0,ii))) Else DV=UserDefineFieldArr(4,ii)
			       KS.Echo  KSUser.GetSelectOption(ChannelID,UserDefineFieldValueStr,UserDefineFieldArr,UserDefineFieldArr(3,II),UserDefineFieldArr(0,ii),UserDefineFieldArr(7,ii),UserDefineFieldArr(5,ii),DV)
				 Exit For
			   End If
			 Next
		   elseif instr(sTemp,"|radio")<>0 then  '��ѡ
			 For II=0 To Ubound(UserDefineFieldArr,2)
			   If Lcase(UserDefineFieldArr(0,Ii))=Lcase(split(sTemp,"|")(0)) Then
			      If Action="Edit" Then DV=RS(Trim(UserDefineFieldArr(0,ii))) Else DV=UserDefineFieldArr(4,ii)
			       KS.Echo  KSUser.GetRadioOption(UserDefineFieldArr(0,ii),UserDefineFieldArr(5,ii),DV)
				 Exit For
			   End If
			 Next
		   elseif instr(sTemp,"|checkbox")<>0 then  '��ѡ
			 For II=0 To Ubound(UserDefineFieldArr,2)
			   If Lcase(UserDefineFieldArr(0,Ii))=Lcase(split(sTemp,"|")(0)) Then
			      If Action="Edit" Then DV=RS(Trim(UserDefineFieldArr(0,ii))) Else DV=UserDefineFieldArr(4,ii)
			       KS.Echo  KSUser.GetCheckBoxOption(UserDefineFieldArr(0,ii),UserDefineFieldArr(5,ii),DV)
				 Exit For
			   End If
			 Next
		   elseif instr(sTemp,"|unit")<>0 then  '��λ
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
		   if err.number<>0 then ks.die "<script>alert('�ֶ�û���ҵ�,�����̨ģ�͹���->�޸�->Ͷ��ѡ���¼��ģ��!');</script>" : err.clear
	end select
	Parse    = iPosBegin
 End Function
 
 
'=========================ɨ���Ա���������� ������2010��6��========================================

Public Sub Kesion()
         Dim LoginTF:LoginTF=Cbool(KSUser.UserLoginChecked)

        '==========================��������״̬================================
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
		 FCls.RefreshType = "Member"   '���õ�ǰλ��Ϊ��Ա����
		 Set KSR = New Refresh
		 TemplateFile=KS.Setting(116)
		 If LoginTF=True Then  TemplateFile=KS.U_G(KSUser.GroupID,"templatefile")
		 If trim(TemplateFile)="" Then TemplateFile=KS.Setting(116)
         If trim(TemplateFile)="" Then Response.Write "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!":response.end
		 FileContent = KSR.LoadTemplate(TemplateFile)
		 If Trim(FileContent) = "" Then FileContent = "ģ�岻����!"
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
			   KS.Echo "<img src=""../images/face/boy.jpg"" width=""62"" height=""58"" alt=""��������"" />"
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
				 KS.Echo ("<li>�ǳƣ�" & KSUser.UserName & "</li><li>���<span style=""curson:pointer"" title=""" & KS.U_G(KSUser.GroupID,"groupname") & """>" & KS.Gottopic(KS.U_G(KSUser.GroupID,"groupname"),12) & "</span></li>")
				 %>
				 <li>״̬��<span class="rl" style="cursor:pointer" onMouseover="$('#userstatus').show();"><%if KSUser.GetUserInfo("IsOnline")="1" then response.write "<img src=""images/online.gif"" align='absmiddle'> ����" else response.write "<img src=""images/notonline.gif"" align='absmiddle'> ����"%>
										<div id="userstatus" class='abs' style="top:15px;width:100px" onMouseOut="$('#userstatus').hide();">
											<dl><img src="images/downline.png" align="absmiddle" />  <a href="?action=offline">��������</a></dl>
											<dl><img src="images/online.png" align="absmiddle" /> <a href="?action=setonline">��������</a></dl>
										</div>
										</span>
				 </li>
				 <li><a href="User_EditInfo.asp?Action=face" style="color:#0066CC;text-decoration:underline">�޸�ͷ��</a> <a href="user_editinfo.asp" style="color:#0066CC;text-decoration:underline">�û�����</a></li>
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
			if KSUser.GetUserInfo("Sex")="��" then
			 KS.Echo "����</span></strong> "
			Else
			 KS.Echo "Ůʿ</span></strong> "
			End If
			If (Hour(Now) < 6) Then
            Str= "�賿�ã�"
			ElseIf (Hour(Now) < 9) Then
			Str= "���Ϻã�"
			ElseIf (Hour(Now) < 12) Then
			Str= "����ã�"
			ElseIf (Hour(Now) < 14) Then
			Str= "����ã�"
			ElseIf (Hour(Now) < 17) Then
			Str= "����ã�"
			ElseIf (Hour(Now) < 18) Then
			Str= "����ã�"
			Else
			Str= "���Ϻã�"
			End If
			KS.Echo str & " ��ӭ������Ա���ģ�<br/>"
			
			
	  If KS.SSetting(0)<>0 Then  '�ж���û�п�ͨ�ռ�
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
		   KS.Echo ("�ҵĿռ��ַ:<a href='" & spacedomain & "' target='_blank'>" & spacedomain & "</a>&nbsp;&nbsp; <span  class='rl' onMouseOver=""$('#myspace').show();"" onMouseOut=""$('#myspace').hide();"" style=""color:#ff6600;""><span style='font-weight:bold;font-size:14px'>�ռ�����</span><img src='images/dico.gif' align='absmiddle'/> <div class='abs' style='width:100px;padding-top:-10px;top:16px;' id='myspace'>")
		    If Conn.Execute("Select top 1 ID From ks_EnterPrise Where UserName='" &KSUser.UserName & "'").eof Then
			   KS.Echo ("<dl> <a href='User_Blog.asp?Action=BlogEdit'><img src='images/menu_icon.gif'> ���˿ռ�����</a></dl>")
			   KS.Echo ("<dl> <a href='user_Enterprise.asp'><img src='images/menu_icon.gif'> ����Ϊ��ҵ�ռ�</a></dl>")
			Else
			   KS.Echo ("<dl> <a href='User_Blog.asp?Action=BlogEdit'><img src='images/menu_icon.gif'> ��ҵ�ռ�����</a></dl>")
			   KS.Echo ("<dl> <a href='User_Blog.asp?action=Template'><img src='images/menu_icon.gif'> �ռ�ģ���</a></dl>")
			End If

			
		   KS.Echo "</div></span>"
		 ELSEIf KSUser.GetUserInfo("usertype")="1" Then
		    KS.Echo "<a href='../company/show.asp?username=" & ksuser.username & "' target='_blank'>" & KS.GetDomain & "company/show.asp?username=" & KSUser.UserName &"</a>"
		 End If
	End If
 End Sub
 
 Sub ShowMyMenu()
   If KS.SSetting(0)=1 Then  '��ͨ�ռ����˳�
		 If KSUser.CheckPower("s02")=true And KSUser.CheckPower("s01")=true then
			 Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			 Response.Write "<img src=""images/icon7.png"" align=""absmiddle"" /> <a href=""User_Blog.asp"">����</a>"
			 Response.Write "<span><a href=""User_Blog.asp?Action=Add"">+����</a></span></li>"
		 End If
		 If KSUser.CheckPower("s05")=true And KSUser.CheckPower("s01")=true Then
			 Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			 Response.Write "<img src=""images/icon2.png"" align=""absmiddle"" /> <a href=""User_Photo.asp"">���</a>"
			 Response.Write "<span><a href=""User_Photo.asp?Action=Add"">+�ϴ�</a></span></li>"
		 End If
		 If KSUser.CheckPower("s06")=true And KSUser.CheckPower("s01")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon19.png"" align=""absmiddle"" /> <a href=""User_Team.asp"">Ȧ��</a>"
			Response.Write "<span><a href=""User_Team.asp?action=CreateTeam"">+����</a></span></li>"
		 End If

	  If Conn.Execute("Select top 1 ID From ks_EnterPrise Where UserName='" &KSUser.UserName & "'").eof Then '���˿ռ�
		 If KSUser.CheckPower("s04")=true And KSUser.CheckPower("s01")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon20.png"" align=""absmiddle"" /> <a href=""User_music.asp"">����</a>"
			Response.Write "<span><a href=""User_Music.asp?action=addlink"">+���</a></span></li>"
		 End If
		 If KSUser.CheckPower("s10")=true Then 
			If KS.C_S(5,21)="1" Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon6.png"" align=""absmiddle"" /> <a href=""user_myshop.asp"">��Ʒ</a>"
			Response.Write "<span><a href=""user_myshop.asp?ChannelID=5&Action=Add"">+����</a></span></li>"
			End If
		 End IF
	  Else   '��ҵ�ռ�
	      if KSUser.CheckPower("s10")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon6.png"" align=""absmiddle"" /> <a title='��ҵ��Ʒ����' href=""user_myshop.asp"">��Ʒ</a>"
			Response.Write "<span><a href=""user_myshop.asp?ChannelID=5&Action=Add"">+����</a></span></li>"
		  End If
		  if KSUser.CheckPower("s11")=true And KSUser.CheckPower("s01")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon21.png"" align=""absmiddle"" /> <a title='��ҵ���Ź���' href=""user_EnterpriseNews.asp"">��̬</a>"
			Response.Write "<span><a href=""user_EnterpriseNews.asp?Action=Add"" title='������ҵ����'>+����</a></span></li>"
		  end if
		  if KSUser.CheckPower("s12")=true And KSUser.CheckPower("s01")=true then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a title='�ؼ��ʹ��������' href=""user_EnterpriseAD.asp"">���</a>"
			Response.Write "<span><a href=""user_EnterpriseAD.asp?Action=Add"">+����</a></span></li>"
		  end if
		  if KSUser.CheckPower("s13")=true And KSUser.CheckPower("s01")=true then
			  	Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
				Response.Write "<img src=""images/icon22.png"" align=""absmiddle"" /> <a title='��ҵ����֤�����' href=""user_Enterprisezs.asp"">����</a>"
				Response.Write "<span><a href=""user_Enterprisezs.asp?Action=Add"">+����</a></span></li>"
		  End If
	  End If
  End If
   		 If KSUser.CheckPower("s03")=true Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon18.png"" align=""absmiddle"" /> <a href=""User_friend.asp"">����</a>"
			Response.Write "<span><a href=""User_Friend.asp?action=addF"">+Ѱ��</a></span></li>"
		 End If

  
  
  	 
'ģ�͵�Ͷ��
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
					Response.Write "<span><a href=""User_MyExam.asp?action=record"">+��¼</a></span></li>"
					Else
					Response.Write "<span><a href=""" & ItemUrl &"?channelid="& Node.SelectSingleNode("@ks0").text & "&Action=Add"">+����</a></span></li>"
					End If
			 Next
	   End If
		 
		 '��ְ
		If KS.C_S(10,21)=1 Then
			If KSUser.GetUserInfo("UserType")=0 Then
				If KSUser.CheckPower("s14")=true Then 
					 If KS.C_S(10,21)="1" Then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a href=""User_JobResume.asp"">�ҹ���</a>"
						Response.Write "<span><a href=""User_JobResume.asp"">+����</a></span></li>"

					 End If
				End If
			Else			 
			if KSUser.CheckPower("s14")=true  then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a href=""user_Enterprise.asp?action=job"">���˲�</a>"
						Response.Write "<span><a href=""User_JobCompanyZW.asp?Action=Add"">+����</a></span></li>"
			 end if
            End If
		End If
		
		 If KS.C_S(5,21)=1 Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon9.png"" align=""absmiddle"" /> <a href=""user_order.asp"">����</a>"
			Response.Write "<span><a href=""user_order.asp?action=coupon"">�Ż�ȯ</a></span></li>"
		 End if
		
		 If KSUser.CheckPower("s07")=true Then 
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon7.png"" align=""absmiddle"" /> <a href=""User_Class.asp"">ר��</a>"
			Response.Write "<span><a href=""User_Class.asp?Action=Add"">+����</a></span></li>"
		 End If
		 If KSUser.CheckPower("s09")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon11.png"" align=""absmiddle"" /> <a href=""User_Askquestion.asp"">�ʴ�</a>"
			Response.Write "<span><a  href=""../ask/a.asp"" target=""_blank"">+����</a></span></li>"
		 End If
		 If KSUser.CheckPower("s19")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon11.png"" align=""absmiddle"" /> <a href=""User_mytopic.asp"">��̳</a>"
			Response.Write "<span><a  href=""User_mytopic.asp?action=fav"">�ղ���</a></span></li>"
		 End If
		 If KSUser.CheckPower("s20")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon13.png"" align=""absmiddle"" /> <a href=""User_ItemSign.asp"">ǩ��</a>"
			Response.Write "<span><a  href=""User_ItemSign.asp"">�鿴</a></span></li>"
		 End If
         if KSUser.CheckPower("s16")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon13.png"" align=""absmiddle"" /> <a href=""User_favorite.asp"">�ղ�</a>"
			Response.Write "<span><a  href=""User_favorite.asp"">�鿴</a></span></li>"
		 End If
         if KSUser.CheckPower("s17")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon15.png"" align=""absmiddle"" /> <a href=""user_feedback.asp"">Ͷ��</a>"
			Response.Write "<span><a  href=""user_feedback.asp?Action=Add"">+����</a></span></li>"
		 End If
End Sub
 
'------ɨ���Ա����������------

 
 
 
 'ȡ��ĳ���ֶε�Ĭ��ֵ
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

'���� isTemplate true ��̨���ɱ�ģ�����,channelid ģ��id, id �༭ʱ������ID
Function GetInputForm(IsTemplate,ChannelID,id,KSUser,RS)
  Dim F_B_Arr,F_V_Arr,UserDefineFieldValueStr
  Dim ClassID,Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,Intro,FullTitle,ReadPoint,Province,City,UserDefineFieldArr,I,SelButton,MapMarker
  F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
  F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
  
if IsObject(RS) And IsTemplate=false Then
	If Not RS.Eof Then
		     If KS.C_S(ChannelID,42) =0 And RS("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
			   RS.Close():Set RS=Nothing
			   KS.ShowTips "error",server.urlencode("��Ƶ�����������" & KS.C_S(ChannelID,3) & "�������޸�!")
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
			 SelButton="ѡ����Ŀ..."
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
			  response.write "�޸�" & KS.C_S(ChannelID,3)
		  Else
		      response.write "����" & KS.C_S(ChannelID,3)
		 End iF%></td>
 </tr>
 <tr class="tdbg">
  <td width="12%"  height="25" align="center"><span><%=F_V_Arr(1)%>��</span></td>
  <td width="88%"><%
				If IsTemplate Then
				  Response.Write "[#SelectClassID]"
				Else
				 Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) 
				End If
			 %></td>
 </tr>
 <tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(0)%>��</span></td>
    <td><input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" /><span style="color: #FF0000">*</span></td>
 </tr>
<%if F_B_Arr(2)=1 Then%> <tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(2)%>��</span></td>
    <td><input class="textbox" name="FullTitle" type="text" style="width:250px; " value="<%=FullTitle%>" maxlength="100" /><span class="msgtips"> �������⣬������</span></td>
 </tr>
<%End If%>
<%if F_B_Arr(5)=1 Then%> <tr class="tdbg">
    <td height="25" align="center"><span><%=F_V_Arr(5)%>��</span></td>
    <td><input name="KeyWords"  class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:220px; " /><a href="javascript:void(0)" onclick="GetKeyTags()" style="color:#ff6600">���Զ���ȡ��</a> <span class="msgtips">����ؼ�������Ӣ�Ķ���(&quot;<span style="color: #FF0000">,</span>&quot;)����</span></td>
  </tr>
<%end if%>
<%if F_B_Arr(6)=1 Then%> <tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(6)%>��</span></td>
    <td height="25"><input name="Author" class="textbox" type="text" id="Author" style="width:220px; " value="<%=Author%>" maxlength="30" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>������<span></td>
  </tr>
<%end if%>
<%if F_B_Arr(7)=1 Then%>  <tr class="tdbg">
   <td  height="25" align="center"><span><%=F_V_Arr(7)%>��</span></td>
   <td><input class="textbox" name="Origin" type="text" id="Origin" style="width:220px; " value="<%=Origin%>" maxlength="100" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>����Դ<span></td>
  </tr>
<%end if%>
<%if F_B_Arr(23)="1" Then%>	<tr class="tdbg">
    <td  height="25" align="center"><span><%=F_V_Arr(23)%>��</span></td>
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
    <td height="25" align="center"><span><%=F_V_Arr(25)%>��</span></td>
    <td>��γ�ȣ�<input value="<%=MapMarker%>" type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='images/edit_add.gif' align='absmiddle' border='0'>��ӵ��ӵ�ͼ��־</a>
	 <script type="text/javascript">
		  function addMap(){
		  new KesionPopup().PopupCenterIframe('���ӵ�ͼ��ע','../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),760,430,'auto');
		  }
		  </script>
	</td>
  </tr>
<%end if%>
<%Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)%>
<%if F_B_Arr(8)=1 Then%> <tr class="tdbg">
   <td  height="25" align="center"><span><%=F_V_Arr(8)%>��</span><br><input name='AutoIntro' type='checkbox' checked value='1'><font color="#FF0000">�Զ���ȡ���ݵ�200������Ϊ����</font></td>
   <td><textarea class='textarea' name="Intro" style='width:95%;height:95px'><%=intro%></textarea></td>
  </tr>
<%end if%>
<%if F_B_Arr(9)=1 Then%> <tr class="tdbg">
   <td><%=F_V_Arr(9)%>:<br><img src="images/ico.gif" width="17" height="12" /><font color="#FF0000">���<%=KS.C_S(ChannelID,3)%>�ϳ�����ʹ�÷�ҳ��ǩ��[NextPage]</font></td>
   <td><%
								If F_B_Arr(21)=1 Then
	%>
			      <table border='0' width='100%' cellspacing='0' cellpadding='0'>
			       <tr><td height='35' width=70>&nbsp;<strong><%=F_V_Arr(21)%>:</strong></td><td width="170"><input onclick='PopInsertAnnex()' class='button' type='button' name='annexbtn' value='�����Ѵ��ڵĸ���'/>&nbsp;</td><td><iframe id='UpFileFrame' name='UpFileFrame' src='User_UpFile.asp?Type=File&ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='200' height='24'></iframe></td></tr>
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
    <td height="25" align="center"><%=F_V_Arr(10)%>��</td>
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
        <td height="25" align="center"><span>�Ķ�<%=KS.Setting(45)%>��</span></td>
         <td height="25"><input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6"><%=KS.Setting(46)%> <span class="msgtips">�������Ķ������롰<font color=red>0</font>��</span></td>
   </tr>
<%end if%>


 <tr class="tdbg">
   <td height="40"></td><td><button class="pn" id="submit1" type="submit"><strong>OK, �� ��</strong></button>&nbsp;<%if IsTemplate Then 
   Response.Write "[#Status]" 
   Elseif id<>0 Then
		     If Verific<>1 Then
			  if Verific=2 Then
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""ȷ������Ͷ�巢����?""));}' checked>����ݸ�</label>"
			  else
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""ȷ������Ͷ�巢����?""));}' value='2'>����ݸ�</label>"
			  end if
			 Else
			  response.write "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
	Else
		    response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2'>����ݸ�</label>"
   End If%> </td>
 </tr>
</table>
<br/><%
End Function
%>