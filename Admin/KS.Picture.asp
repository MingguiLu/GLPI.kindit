<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Picture
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Picture
        Private KS,KSCls
		'=====================================���屾ҳ��ȫ�ֱ���=====================================
		Private ID, I, totalPut, Page, RS,ComeFrom
		Private KeyWord, SearchType, StartDate, EndDate, ParentRs, SearchParam,MaxPerPage,SpecialID
		Private T, TitleStr, AttributeStr
		Private FolderID, TemplateID,WapTemplateID,TN, TI,TJ,Action,UserDefineFieldArr,UserDefineFieldValueStr
		Private PicID, Title, PhotoUrl, PictureContent, PicUrls, Recommend,IsTop
		Private Popular, Strip, Verific, Comment, Slide, ChangesUrl, Rolls, KeyWords, Author, Origin, AddDate, Rank, Hits, HitsByDay, HitsByWeek, HitsByMonth
		Private CurrPath, InstallDir,PreViewObj, UpPowerFlag,Inputer,SaveFilePath,MapMarker
		Private ComeUrl,F_B_Arr,F_V_Arr,ChannelID,FileName,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj,Score,ShowStyle,PageNum
		Private ReadPoint,ChargeType,PitchTime,ReadTimes,InfoPurview,arrGroupID,DividePercent
		'=============================================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
        F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
		
		'�ռ���������
		KeyWord   = KS.G("KeyWord")
		SearchType= KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate   = KS.G("EndDate")
		Action     = KS.G("Action")
		ComeFrom   = KS.G("ComeFrom")
		SearchParam = "ChannelID=" & ChannelID
		If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
		If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
		If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
		If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
		If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
		If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom
		
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))

			Action = Trim(KS.G("Action"))
			Page = KS.G("page")
							
			IF KS.G("Method")="Save" Then
				 Call PictureSave()
			Else 
				 Call PictureAdd()
			End If
		End Sub

        '���
        Sub PictureAdd() 
			With Response
			CurrPath = KS.GetUpFilesDir()
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Action = "Add" Then
			  FolderID = Trim(KS.G("FolderID"))
			  
			  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10002") Then          '����Ƿ������ͼƬ��Ȩ��
			   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&ChannelID=" & ChannelID &"';</script>")
			   Call KS.ReturnErr(1, "")
			   Exit Sub
			  End If
			  Hits = 0:HitsByDay = 0: HitsByWeek = 0:HitsByMonth = 0:Comment = 1:IsTop=0:UserDefineFieldValueStr=0
			  ReadPoint=0:PitchTime=24:ReadTimes=10:Score=0 : ShowStyle=4: PageNum=12
			  PreViewObj = "<br><br><br>" & KS.C_S(ChannelID,3) & "Ԥ����"
			  KeyWords = Session("keywords")
			  Author = Session("Author")
			  Origin = Session("Origin")
			
			ElseIf Action = "Edit" Or Action="Verify" Then

			   Set RS = Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & KS.ChkClng(KS.G("ID")), conn, 1, 1
			   If RS.EOF And RS.BOF Then
				Call KS.Alert("�������ݳ���!", ComeUrl)
				Set KS = Nothing:.End:Exit Sub
			   End If
				PicID = Trim(RS("ID"))
				FolderID = Trim(RS("Tid"))
				If Action ="Edit" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then     '����Ƿ��б༭ͼƬ��Ȩ��
				 RS.Close:Set RS = Nothing
				 If KeyWord = "" Then
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
				 Else
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=" &server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>����" & KS.C_S(ChannelID,3) & "���</font>") & "&ButtonSymbol=PictureSearch';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate)
				 End If
				 Exit Sub
			   End If
			   IF Action="Verify" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10012") Then 
			     RS.Close:Set RS = Nothing
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid & "';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
				 
				 Exit Sub   
			   End If
			   
				Title    = Trim(RS("title"))
				PhotoUrl = Trim(RS("PhotoUrl"))
				PreViewObj = "<img src='" & PhotoUrl & "' border='0'>"
				PicUrls  = Trim(RS("PicUrls"))
				PictureContent = Trim(RS("PictureContent")) : If KS.IsNul(PictureContent) Then PictureContent=" "
				Rolls    = CInt(RS("Rolls"))
				Strip    = CInt(RS("Strip"))
				Recommend = CInt(RS("Recommend"))
				Popular  = CInt(RS("Popular"))
				Verific  = CInt(RS("Verific"))
				Comment  = CInt(RS("Comment"))
				IsTop    = (RS("IsTop"))
				Slide    = CInt(RS("Slide"))
				AddDate  = CDate(RS("AddDate"))
				Rank     = Trim(RS("Rank"))
				FileName = RS("Fname")
				
				TemplateID = RS("TemplateID")
				WapTemplateID=RS("WapTemplateID")
				Hits = Trim(RS("Hits"))
				HitsByDay = Trim(RS("HitsByDay"))
				HitsByWeek = Trim(RS("HitsByWeek"))
				HitsByMonth = Trim(RS("HitsByMonth"))
				Score=RS("Score")
				KeyWords = Trim(RS("KeyWords"))
				Author = Trim(RS("Author"))
				Origin = Trim(RS("Origin"))
				FolderID = RS("Tid")
				ShowStyle= RS("ShowStyle")
				PageNum=RS("PageNum")
				ReadPoint = RS("ReadPoint")
				ChargeType= RS("ChargeType")
				PitchTime = RS("PitchTime")
				ReadTimes = RS("ReadTimes")
				InfoPurview=RS("InfoPurview")
				arrGroupID = RS("arrGroupID")
				DividePercent=RS("DividePercent")
				If F_B_Arr(18)="1" Then	MapMarker      = RS("MapMarker")
               '�Զ����ֶ�
				UserDefineFieldArr=KSCls.Get_KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				  Dim UnitOption
				  If UserDefineFieldArr(12,I)="1" Then
				   UnitOption="@" & RS(UserDefineFieldArr(0,I)&"_Unit")
				  Else
				   UnitOption=""
				  End If
				  If I=0 Then
				    UserDefineFieldValueStr=RS(UserDefineFieldArr(0,I)) &UnitOption & "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & RS(UserDefineFieldArr(0,I)) &UnitOption& "||||"
				  End If
				Next
			  End If
				RS.Close
			End If
			'ȡ���ϴ�Ȩ��
			UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
			
           ' .Write"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">"
			.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrlf
			.Write "<title>���</title>" & vbCrlf
			.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>" & vbCrlf
			.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>" & vbCrlf
			.Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>" & vbCrlf
			.Write "<script language=""javascript"" src=""../KS_Inc/popcalendar.js""></script>" & vbCrlf
			.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
			.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
			.Write "<script type=""text/javascript"" src=""../editor/ckeditor.js"" mce_src=""../editor/ckeditor.js""></script>"
			.Write "<script language='javascript' src='../ks_inc/kesion.box.js'></script>"

			.Write "</head>" & vbCrlf
			.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>" & vbCrlf
			.Write "<div align='center'>" & vbCrlf
			.Write "<ul id='menu_top'>"
			.Write "<li onclick=""return(SubmitFun())"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>ȷ������</span></li>"
			.Write "<li onclick=""history.back();"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>ȡ������</span></li>"
		    .Write "</ul>" & vbCrlf
			
			.Write "<div class=tab-page id=PhotoPane>"
			.Write " <SCRIPT type=text/javascript>"
			.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""PhotoPane"" ), 1 )"
			.Write " </SCRIPT>"
				 
			.Write " <div class=tab-page id=basic-page>"
			.Write "  <H2 class=tab>������Ϣ</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"
			.Write "	</SCRIPT>"		
			.Write "    <form action='?ChannelID=" & ChannelID & "&Method=Save' method='post' id='myform' name='myform' onsubmit='return(SubmitFun())'>"
            .Write " <TABLE width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			.Write "      <input type='hidden' value='" & PicID & "' name='PicID'>"
			.Write "      <input type='hidden' value='" & Action & "' name='Action'>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
			.Write "      <input type='hidden' name='Inputer' value='" &Inputer & "'>"
			
			.Write "       <tr class='tdbg'>"
			.Write "          <td height='20' width='85' class='clefttitle'><div align='right'><font color='#FF0000'><strong>" & F_V_Arr(0) & ":</strong></font></div></td>"
			.Write "          <td height='25' nowrap> "
			.Write "            <input name='title' type='text'  class='textbox' value='" & Title & "' size=80>"
			.Write "                  <font color='#FF0000'>*</font>"
			If F_B_Arr(17)=1 Then
			.Write "<input type='checkbox' name='MakeHtml' value='1' checked>" & F_V_Arr(17)
			End IF
			.Write "                  </td>"
			.Write "       </tr>"
			.Write "       <tr class='tdbg'>"
			.Write "         <td width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(1) & ":</strong></div></td>"
			.Write "         <td><input type='hidden' name='OldClassID' value='"& FolderID & "'>"
			.Write " <select size='1' name='tid' id='tid'>"
			.Write " <option value='0'>--��ѡ����Ŀ--</option>"
			.Write Replace(KS.LoadClassOption(ChannelID,true),"value='" & FolderID & "'","value='" & FolderID &"' selected") & " </select>"

		
		 If F_B_Arr(5)=1 Then
			.Write "&nbsp;&nbsp;" & F_V_Arr(5) & " <input name='Recommend' type='checkbox' id='Recommend' value='1'"
			If Recommend = 1 Then .Write (" Checked")
			.Write ">�Ƽ�"
			.Write "<input name='Rolls' type='checkbox' id='Rolls' value='1'"
			If Rolls = 1 Then .Write (" Checked")
			.Write ">����"
			.Write "<input name='Strip' type='checkbox' id='Strip' value='1'"
			If Strip = 1 Then .Write (" Checked")
			.Write ">ͷ��"
			.Write "<input name='Popular' type='checkbox' id='Popular' value='1'"
			If Popular = 1 Then .Write (" Checked")
			.Write ">����"
			.Write "<input name='IsTop' type='checkbox' id='IsTop' value='1'"
			If IsTop = 1 Then .Write (" Checked")
			.Write ">�̶�"
			.Write "<input name='Comment' type='checkbox' id='Comment' value='1'"
			If Comment = 1 Then .Write (" Checked")
			.Write ">��������"
			.Write "<input name='Slide' type='checkbox' id='Slide' value='1'"
			If Slide = 1 Then
			.Write (" Checked")
			End If
			.Write ">�õ�"
			.Write "</td>"
			.Write "              </tr>"
		 End If
		 
		   If F_B_Arr(18)="1" Then	
		%>
		  <script type="text/javascript">
		  function addMap(){
		  new KesionPopup().PopupCenterIframe('���ӵ�ͼ��ע','../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),760,430,'auto');
		  }
		  </script>
		<%
			.Write "              <tr  class='tdbg' style='height:25px'>"
			.Write "                <td height='25' class='clefttitle'><div align=right><strong>���ӵ�ͼ:</strong></div></td>"
			.Write "                <td height='25' align='left'>��γ�ȣ�<input value=""" & MapMarker & """ type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='images/accept.gif' align='absmiddle' border='0'>��ӵ��ӵ�ͼ��־</a>"
			.Write "              </td>"
			.Write "              </tr>"
		 End If
		 If F_B_Arr(6)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(6) & ":</strong></div></td>"
			.Write "                <td> <input name='KeyWords' type='text' id='KeyWords' class='textbox' value='" & KeyWords & "' size=40> <="
			.Write "                  <select name='SelKeyWords' style='width:150px' onChange='InsertKeyWords($(""#KeyWords"")[0],this.options[this.selectedIndex].value)'>"
		    .Write "<option value="""" selected> </option><option value=""Clean"" style=""color:red"">���</option>"
			.Write KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")
			.Write "                  </select>"
			.Write " <br />��<a href=""#"" id=""KeyLinkByTitle"" style=""color:green"">����" & F_V_Arr(0) & "�Զ���ȡTags</a>��<input type='checkbox' name='tagstf' value='1' checked>д��Tags��</td>"
			.Write "              </tr>"
		End If
		If F_B_Arr(7)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(7) & ":</strong></div></td>"
			.Write "                <td> <input name='author' type='text' id='author' value='" & Author & "' size=30 class='textbox'>                 <=��<font color='blue'><font color='#993300' onclick='$(""#author"").val(""δ֪"");' style='cursor:pointer;'>δ֪</font></font>����<font color='blue'><font color='#993300' onclick=""$('#author').val('����');"" style='cursor:pointer;'>����</font></font>����<font color='blue'><font color='red' onclick=""$('#author').val('" & KS.C("AdminName") & "');"" style='cursor:pointer;'>" & KS.C("AdminName") & "</font></font>��"
							 If Author <> "" And Author <> "δ֪" And Author <> KS.C("AdminName") And Author <> "����" Then
							  .Write ("��<font color='blue'><font color='#993300' onclick=""$(""#author"").val('" & Author & "');"" style='cursor:pointer;'>" & Author & "</font></font>��")
							 End If
							  .Write ("<select name='SelAuthor' style='width:100px' onChange=""$('#author').val(this.options[this.selectedIndex].value);"">")
		    .Write "<option value="""" selected> </option><option value="""" style=""color:red"">���</option>"
			.Write KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=0 And OriginType=1 Order BY AddDate Desc")
			.Write "                                   </select></td>"
			.Write "              </tr>"
	  End If
	  If F_B_Arr(8)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td  width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(8) & ":</strong></div></td>"
			.Write "                <td> <input name='Origin' type='text' id='Origin' value='" & Origin & "' size=30 class='textbox'>                 <=��<font color='blue'><font color='#993300' onclick=""$('#Origin').val('����');"" style='cursor:pointer;'>����</font></font>����<font color='blue'><font color='#993300' onclick=""$('#Origin').val('��վԭ��');"" style='cursor:pointer;'>��վԭ��</font></font>����<font color='blue'><font color='#993300' onclick=""$('#Origin').val('������');"" style='cursor:pointer;'>������</font></font>��"
							  If Origin <> "" And Origin <> "����" And Origin <> "��վԭ��" And Origin <> "������" Then
							  .Write ("��<font color='blue'><font color='#993300' onclick=""$('#Origin').val('" & Origin & "')"" style='cursor:pointer;'>" & Origin & "</font></font>�� ")
							   End If
							  .Write ("<select name='selOrigin' style='width:100px' onChange=""$('#Origin').val(this.options[this.selectedIndex].value)"">")
		    .Write "<option value="""" selected> </option><option value="""" style=""color:red"">���</option>"
			.Write KSCls.Get_O_F_D("KS_Origin","OriginName","OriginType=0 Order BY AddDate Desc")
			.Write "                </select> </td>"
			.Write "              </tr>"
	 End If
			        '�Զ����ֶ�
		    .Write KSCls.Get_KS_D_F(ChannelID,UserDefineFieldValueStr)
			
			.Write "              <tr  class='tdbg' id='mode1' style='height:25px'>"
			.Write "                <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(2) & ":</strong></div></td>"
			.Write "                <td> <input name='PhotoUrl' type='text' id='PhotoUrl' size='30' value='" & PhotoUrl & "' class='textbox'>"
			.Write "   <font color='#FF0000'>*</font>&nbsp;<input class='button' type='button' name='Submit' value='ѡ��ͼƬ...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID & "&CurrPath=" & CurrPath & "',550,290,window,$('#PhotoUrl')[0]);""> <input class='button' type='button' name='Submit' value='Զ��ץͼ...' onClick=""OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('ץȡԶ��ͼƬ')+'&ItemName=ͼƬ&CurrPath=" & CurrPath & "',300,100,window,$('#PhotoUrl')[0]);"">"
			.Write "                  <input class=""button""  type='button' name='Submit' value='�ü�...' onClick=""if($('#PhotoUrl').val()==''){alert('��ѡ��ͼƬ�����ϴ�����ʹ�ô˹���');return false;}else{OpenImgCutWindow(1,'" & KS.Setting(3) & "',$('#PhotoUrl').val())}"">"
			
			If Action="Add" Then
			.Write "<br/><label><input type='checkbox' name='autothumb' id='autothumb' value='1' checked>ʹ��ͼ���ĵ�һ��ͼ</label>"
			End If
			
			.Write "     </td>"
			.Write "              </tr>"
			
			.Write "              <tr  class='tdbg'>"
			.Write "                <td class='clefttitle'><div align=right><strong>�ϴ�ͼƬ:</strong></div></td>"
			.Write "                <td  align='left'><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?UPType=Pic&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='40'></iframe>"
			.Write "              </td>"
			.Write "              </tr>"
			
			 
			.Write "<tr class='tdbg'><td height='25' nowrap align='right' class='clefttitle'><strong>��ʾ��ʽ:</strong></td><td><table width='80%'><tr><td><input type='radio' onclick=""$('#pagenums').hide();"" name='showstyle' value='4'"
			If ShowStyle="4" Then .Write " checked"
			.Write "><img src='../images/default/p4.gif' title='��ͼƬ��ֻ��һ��ͼƬʱ��Ч,���ô���ʽ��Ч!'></td><td><input type='radio' onclick=""$('#pagenums').hide();"" name='showstyle' value='1'"
			If ShowStyle="1" Then .Write " checked"
			.Write "><img src='../images/default/p1.gif' title='��ͼƬ��ֻ��һ��ͼƬʱ��Ч,���ô���ʽ��Ч!'></td>"
			.Write "<td><input type='radio' onclick=""$('#pagenums').show();"" name='showstyle' value='2'"
			If ShowStyle="2" Then .Write " checked"
			.Write "><img src='../images/default/p2.gif' title='��ͼƬ��ֻ��һ��ͼƬʱ��Ч,���ô���ʽ��Ч!'></td>"
			.Write "<td><input type='radio' onclick=""$('#pagenums').show();"" name='showstyle' value='3'"
			If ShowStyle="3" Then .Write " checked"
			.Write"><img src='../images/default/p3.gif'></td>"
			.Write "</tr></table><div id=""pagenums"""
			If ShowStyle="1" or ShowStyle="4" Then .Write " style='display:none'"
			.Write ">ÿҳ��ʾ<input type=""text"" name=""pagenum"" value=""" & PageNum & """ style=""text-align:center;width:30px"">��</div></td></tr>"

            If KS.G("Action")<>"Add" Then
			.Write "      <tr  class='tdbg' style='display:none'>"
			Else
			.Write "      <tr class='tdbg'>"
			End If
			
			.Write "      <td height='25' nowrap align='right' class='clefttitle'><strong>���ģʽ:</strong></td>"
			.Write "      <td>"
			.Write "<label><input type='radio' name='addmode' value='0' checked onclick='$(""#addmore"").hide();$(""#addarea"").show();'>ֱ�����</label> <label><input type='radio' name='addmode' value='1' onclick='$(""#addmore"").show();$(""#addarea"").hide()'>�������</label>"
			.Write "                </td>"
			.Write "              </tr>"
			Dim CurrDate:CurrDate=Year(Now) &right("0"&Month(Now),2)
			Dim CurrDay:CurrDay=CurrDate & right("0"&day(Now),2)
			.Write "            <tr class='tdbg' id='addmore' style='display:none'>"
			.Write "               <td height='35' align='right' class='clefttitle'><strong>" & KS.C_S(ChannelID,3) & "��ַ:</strong></td>"
			.Write "               <td height='25'><input  name='MorePicUrl' type='text' id='MorePicUrl' size='80' value='ͼƬ#|" &  CurrPath & "/"&CurrDate &"/" & CurrDay & "#.jpg|" &  CurrPath & "/"&CurrDate &"/" & CurrDay & "#_S.jpg' class='upfile'><br>&nbsp;&nbsp;��ʼID��<input class='textbox' type='text' value='1' name='morestart' size=5> ����ID��<input class='textbox' type='text' value='100' name='moreend' size=5><font color=red> �������ͨ���Ϊ#��ע��ͨ���ֻ��һ��#����</font><br>&nbsp;&nbsp;<font color=green>��ʽ��ͼƬ����|��ͼ��ַ|Сͼ��ַ</font></td>"
			.Write "              </tr>"


			.Write "<tbody  class='tdbg' id='addarea'>"			
			.Write "<tr class='tdbg'><td width='85' class='clefttitle' align='right' valign='top'><b>" & F_V_Arr(4) & ":</b><br/><br/><span style='color:green;'><input type='checkbox' value='1' name='BeyondSavePic' checked>����ӵ������ϵ�ַʱ,�Զ��ɼ���ͼ<br></span></td><td>"
			%>
			<style type="text/css">
			#thumbnails{background:url(../plus/swfupload/images/albviewbg.gif) no-repeat;min-height:200px;_height:expression(document.body.clientHeight > 200? "200px": "auto" );}
			#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:162px;height:155px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
			#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px}
			</style>
			<link href="../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
			<script type="text/javascript" src="../plus/swfupload/swfupload/swfupload.js"></script>
			<script type="text/javascript" src="../plus/swfupload/js/handlers.js"></script>
<script type="text/javascript">
		var swfu;
		var pid=0;
		function SetAddWater(obj){
		 if (obj.checked){
		 swfu.addPostParam("AddWaterFlag","1");
		 }else{
		 swfu.addPostParam("AddWaterFlag","0");
		 }
        }
		//ɾ���Ѿ��ϴ���ͼƬ
		function DelUpFiles(pid)
		{
		 $("#thumbshow"+pid).remove();
		}	
		
		function addImage(bigsrc,smallsrc,text) {
			var newImgDiv = document.createElement("div");
			var delstr = '';
			delstr = '<a href="javascript:DelUpFiles('+pid+')" style="color:#ff6600">[ɾ��]</a>';
			newImgDiv.className = 'thumbshow';
			newImgDiv.id = 'thumbshow'+pid;
			document.getElementById("thumbnails").appendChild(newImgDiv);
			newImgDiv.innerHTML = '<a href="'+bigsrc+'" target="_blank"><span id="show'+pid+'"></span></a>';
			newImgDiv.innerHTML += '<div style="margin-top:10px;text-align:left">'+delstr+' <b>ע�ͣ�</b><input type="hidden" class="pics" id="pic'+pid+'" value="'+bigsrc+'|'+smallsrc+'"/><input type="text" name="picinfo'+pid+'" value="'+text+'" style="width:155px;" /></div>';
		
			var newImg = document.createElement("img");
			newImg.style.margin = "5px";
		
			document.getElementById("show"+pid).appendChild(newImg);
			if (newImg.filters) {
				try {
					newImg.filters.item("DXImageTransform.Microsoft.Alpha").opacity = 0;
				} catch (e) {
					newImg.style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=' + 0 + ')';
				}
			} else {
				newImg.style.opacity = 0;
			}
		
			newImg.onload = function () {
				fadeIn(newImg, 0);
			};
			newImg.src = smallsrc;
			pid++;
			
		}
	
		window.onload = function () {
			swfu = new SWFUpload({
				// Backend Settings
				upload_url: "include/swfupload.asp",
				post_params: {AddWaterFlag:"1","BasicType":<%=KS.C_S(ChannelID,6)%>,"ChannelID":<%=ChannelID%>,"AutoRename":4},

				// File Upload Settings
				file_size_limit : "2 MB",	// 2MB
				file_types : "*.jpg; *.gif; *.png",
				file_types_description : "֧��.JPG.gif.png��ʽ��ͼƬ,���Զ�ѡ",
				file_upload_limit : 0,

				// Event Handler Settings - these functions as defined in Handlers.js
				//  The handlers are not part of SWFUpload but are part of my website and control how
				//  my website reacts to the SWFUpload events.
				swfupload_preload_handler : preLoad,
				swfupload_load_failed_handler : loadFailed,
				file_queue_error_handler : fileQueueError,
				file_dialog_complete_handler : fileDialogComplete,
				upload_progress_handler : uploadProgress,
				upload_error_handler : uploadError,
				upload_success_handler : uploadSuccess,
				upload_complete_handler : uploadComplete,

				// Button Settings
				//button_image_url : "../plus/swfupload/images/SmallSpyGlassWithTransperancy_17x18d.png",
				button_placeholder_id : "spanButtonPlaceholder",
				button_width: 195,
				button_height: 22,
				button_text : '<span class="button">���������ϴ�(��ͼ����2 MB)</span>',
				button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#666666;font-size: 14px; } ',
				button_text_top_padding: 3,
				button_text_left_padding: 0,
				button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
				button_cursor: SWFUpload.CURSOR.HAND,
				
				// Flash Settings
				flash_url : "../plus/swfupload/swfupload/swfupload.swf",
				flash9_url : "../plus/swfupload/swfupload/swfupload_FP9.swf",

				custom_settings : {
					upload_target : "divFileProgressContainer"
				},
				
				// Debug Settings
				debug: false
			});
		};
	</script>
	<script type="text/javascript">
	function OnlineCollect(){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("<div style='text-align:left;padding-left:2px'><img src='images/folder/R.gif' align='absmiddle'/>���ϲɼ�ͼƬ</div>","<div style='padding:3px'>��http://��ͷ��Զ��ͼƬ��ַ,ÿ��һ��ͼƬ��ַ:<br/><textarea id='collecthttp' style='width:400px;height:150px'></textarea><br/><input type='button' value='ȷ ��' onclick='ProcessCollect()' class='button'/> <input type='button' value='ȡ ��' class='button' onclick='closeWindow()'/></div>",420,340);
	}
	function AddTJ(){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("<div style='text-align:left;padding-left:2px'><img src='images/folder/R.gif' align='absmiddle'/>���ϴ��ļ���ѡ��</div>","<div style='padding:3px'><strong>Сͼ��ַ:</strong><input type='text' name='x1' id='x1'> <input type='button' onclick=\"OpenThenSetValue('Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,$('#x1')[0]);\" value='ѡ��Сͼ' class='button'/><br/><strong>��ͼ��ַ:</strong><input type='text' name='x2' id='x2'> <input type='button' onclick=\"OpenThenSetValue('Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,$('#x2')[0]);\" value='ѡ���ͼ' class='button'/><br/><strong>��Ҫ����:</strong><input type='text' name='x3' id='x3'><br/><br/><input type='button' value='�� ��' onclick='ProcessAddTj()' class='button'/> <input type='button' value='ȡ ��' class='button' onclick='closeWindow()'/></div>",420,340);
	}
	function ProcessAddTj(){
	  if ($("#x1").val()==''){
	   alert('��ѡ��һ��Сͼ��ַ!');
	   $("#x1").focus();
	   return false;
	  }
	  if ($("#x2").val()==''){
	   alert('��ѡ��һ�Ŵ�ͼ��ַ!');
	   $("#x2").focus();
	   return false;
	  }
	  addImage($("#x2").val(),$("#x1").val(),$("#x3").val())
	  $("#x2").val('');
	  $("#x1").val('');
	  $("#x3").val('');
	}
	function ProcessCollect(){
	 var collecthttp=$("#collecthttp").val();
	 if (collecthttp==''){
	   alert('������Զ��ͼƬ��ַ,һ��һ�ŵ�ַ!');
	   $("#collecthttp").focus();
	   return false;
	 }
	 var carr=collecthttp.split('\n');
	 for(var i=0;i<carr.length;i++){
	   
	   var bigsrc=carr[i];
	   var smallsrc=carr[i];
	   addImage(bigsrc,smallsrc,'')
	 }
	 //$("#collecthttp").empty();
	 closeWindow();
	}
	</script>
	    <table>
		 <tr>
		  <td>

	    <div class="pn" style="margin: -6px 0px 0 0;">
		 <span id="spanButtonPlaceholder"></span>
		 		
		</div>
		 </td>
		 <td>
		 <button type="button"  class="pn" onclick="OnlineCollect()" style="margin: -6px 0px 0 0;"><strong>���ϲɼ�</strong></button>
		 <button type="button"  class="pn" onClick="AddTJ();" style="margin: -6px 0px 0 0;"><strong>ͼƬ��...</strong></button>
		 </td>
		 </tr>
		</table>

		<label><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)" checked="checked"/>ͼƬ���ˮӡ</label>
		<div id="divFileProgressContainer"></div>
		
	<div id="thumbnails"></div>
			<input type='hidden' name='PicUrls' id='PicUrls'>
			<%
			.Write "</td></tr>"
			
            .Write "</tbody>"
			
			


		 If F_B_Arr(9)=1 Then
			.Write "      <TR class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(9) & ":</strong></div></td>"
			.Write "                <td nowrap>"
			
			.Write "      <textarea  ID='Content' name='Content' cols=90 rows=6 style='display:none'>" & Server.HTMLEncode(PictureContent) & "</textarea>"

			.Write "<script type=""text/javascript"">"
            .Write "CKEDITOR.replace('Content', {width:""98%"",height:""160px"",toolbar:""Basic"",filebrowserBrowseUrl :""Include/SelectPic.asp?from=ckeditor&Currpath="& KS.GetUpFilesDir() &""",filebrowserWindowWidth:650,filebrowserWindowHeight:290});"
			.Write "</script>"
			
			
			.Write "      </TD></TR>"
		 End If
           .Write "</table>"
		   .Write "</div>"
	
	If F_B_Arr(15)=1 Then		 
		   .Write " <div class=tab-page id=classoption-page>"
		   .Write "  <H2 class=tab>��������</H2>"
		   .Write "	<SCRIPT type=text/javascript>"
		   .Write "				 tabPane1.addTabPage( document.getElementById( ""classoption-page"" ) );"
		   .Write "	</SCRIPT>"

            .Write "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			.Write "           <tr class='tdbg'>"
			.Write "              <td class='clefttitle' width='80' align='right'><strong>����ר��:</strong></td>"
			.Write "              <td>"
			Call KSCls.Get_KS_Admin_Special(ChannelID,KS.ChkClng(KS.G("ID")))
			.write "              </td>"
			.Write "           </tr>"
		 If F_B_Arr(10)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(10) & ":</strong></div></td>"
			.Write "                <td>"
			If Action <> "Edit" Then
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & Now() & "' size='50'  class='textbox'>")
			Else
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & AddDate & "' size='50'  class='textbox'>")
			End If
			.Write "                  <b><a href='#' onClick=""popUpCalendar(this, $('input[name=AddDate]').get(0), dateFormat,-1,-1)""><img src='Images/date.gif' border='0' align='absmiddle' title='ѡ������'></a><strong>���ڸ�ʽ����-��-�� ʱ���֣���</strong>"
			.Write "                  </b></td>"
			.Write "              </tr>"
		End If
		If F_B_Arr(11)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(11) & ":</strong></div></td>"
			.Write "                <td><select name='rank'>"
			If Rank = "��" Then
			.Write "                    <option  selected>��</option>"
			Else
			.Write "                    <option>��</option>"
			End If
			If Rank = "���" Then
			.Write "                    <option  selected>���</option>"
			Else
			.Write "                    <option>���</option>"
			End If
			If Rank = "����" Or Action = "Add" Then
			.Write "                    <option  selected>����</option>"
			Else
			.Write "                    <option>����</option>"
			End If
			If Rank = "�����" Then
			.Write "                    <option  selected>�����</option>"
			Else
			.Write "                    <option>�����</option>"
			End If
			If Rank = "������" Then
			.Write "                    <option  selected>������</option>"
			Else
			.Write "                    <option>������</option>"
			End If
			.Write "                  </select>"
			.Write "                  ��Ϊ" & KS.C_S(ChannelID,3) & "�����Ƽ��ȼ�</td>"
			.Write "              </tr>"
	   End If
	   If F_B_Arr(12)=1 Then
			 .Write "             <tr class='tdbg'>"
			 .Write "               <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(12) & ":</strong></td>"
			 .Write "               <td> ���գ�<input name='HitsByDay' type='text' id='HitsByDay' value='" & HitsByDay & "' size='10' class='textbox'> ���ܣ�<input name='HitsByWeek' type='text' id='HitsByWeek' value='" & HitsByWeek & "' size='10' class='textbox'> ���£�<input name='HitsByMonth' type='text' id='HitsByMonth' value='" & HitsByMonth & "' size='10' class='textbox'> �ܼƣ�<input name='Hits' type='text' id='Hits' value='" & Hits & "' size='10' class='textbox'>"
			 .Write "&nbsp;��Ʊ����<input type='text' name='score' size='6' value='" & score & "'>Ʊ  �����õ�"
			 .Write "             </td>"
			 .Write "             </tr>"
	  End If
	  If F_B_Arr(13)=1 Then
			 .Write "             <tr class='tdbg'>"
			 .Write "               <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(13) & ":</strong></div></td>"
			.Write "                <td> "
			IF Action <> "Edit" and  Action<>"Verify" Then
			.Write " <input type='radio' name='templateflag' onclick='GetTemplateArea(false);' value='2' checked>�̳���Ŀ�趨<input type='radio' onclick='GetTemplateArea(true);' name='templateflag' value='1'>�Զ���"
			.Write "<div id='templatearea' style='display:none'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEBģ��</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>WAPģ��</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			Else
			
			.Write "<div id='templatearea'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEBģ��</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly maxlength='255' size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]")
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>WAPģ��</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			End If
			.Write "                </td>"
			.Write "             </tr>"
	  End If
	  If F_B_Arr(14)=1 Then
			.Write "             <tr class='tdbg'>"
			.Write "               <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(14) & ":</strong></td><td>"
			IF Action = "Edit" or Action="Verify" Then
			.Write "<input name='FileName' type='text' id='FileName' readonly  value='" & FileName & "' size='25' class='textbox'> <font color=red>���ܸ�</font>"
			Else
			.Write "<input type='radio' value='0' name='filetype' onclick='GetFileNameArea(false);' checked>�Զ����� <input type='radio' value='1' name='filetype' onclick='GetFileNameArea(true);' >�Զ���"
			.Write "<div id='filearea' style='display:none'><input name='FileName' type='text' id='FileName'   value='" & FileName  & "' size='25' class='textbox'> <font color=red>�ɴ�·��,�� help.html,news/news_1.shtml��</font></div>"
			End IF
			 .Write "                  </td>"
			 .Write "             </tr>"
	 End If
			
			.Write "</table>"
			.Write "</div>"
  End If
      
	     If F_B_Arr(16)=1 Then
	       KSCls.LoadChargeOption ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent
         End If
		 
	       KSCls.LoadRelativeOption ChannelID,KS.ChkClng(KS.G("ID"))
		   
			 .Write "</form>"
			 .Write " </div>"
			%>
			 <script type="text/javascript">
			 $(document).ready(function(){
				$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
				$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);
			 <%If F_B_Arr(6)=1 Then%>
			  $('#KeyLinkByTitle').click(function(){
			    GetKeyTags();
			  });
			 <%End If%>
			 IniPicUrl();
			});
			function GetKeyTags()
			{
			  var text=escape($('input[name=Title]').val());
			  
			  if (text!=''){
				  $('#KeyWords').val('���Ե�,ϵͳ�����Զ���ȡtags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{
			   alert('�Բ���,������������!');
			  }
			}
			
			function IniPicUrl()
			{
			  var PicUrls='<%=replace(PicUrls,vbcrlf,"\t\n")%>';
			  var PicUrlArr=null;
			  if (PicUrls!='')
			   { 
				PicUrlArr=PicUrls.split('|||');
			    for ( var i=1 ;i<PicUrlArr.length+1;i++){ 
			      addImage(PicUrlArr[i-1].split('|')[1],PicUrlArr[i-1].split('|')[2],PicUrlArr[i-1].split('|')[0]);
			    }
			   }
			}

			function SelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",true);
			  });
			}
			function UnSelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",false);
			  });
			}

			function GetFileNameArea(f)
			{
			  $('#filearea').toggle(f);
			}
			function GetTemplateArea(f)
			{
			   $('#templatearea').toggle(f);
			}
			function SubmitFun()
			{ 	
			    if ($('input[name=title]').val()=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>���ƣ�");
					$('input[name=title]').focus();
					return;
				  }
			   if ($("#tid>option[selected=true]").val()=='0')
			   {
			       alert('��ѡ��������Ŀ!');
				   return false;
			   }
			    
			   
			 	if ($('input[name=PhotoUrl]').val()==''<%if action="Add" Then response.write " && $('#autothumb').attr('checked')==false"%>)
				{
					alert("������<%=KS.C_S(ChannelID,3)%>����ͼ��");
					$('input[name=PhotoUrl]').focus();
					return;
				}
			    
			  var picSrcs='';
			  var src='';
			  $("#thumbnails").find(".pics").each(function(){
			     src=$(this).next().val().replace('|||','').replace('|','')+'|'+$(this).val()
			     if(picSrcs==''){
			      picSrcs=src;
				 }else{
			      picSrcs+='|||'+src;
				 }
			  });
			   $('#PicUrls').val(picSrcs);
	
				var addmode;
				for (var i=0;i<document.myform.addmode.length;i++){
				 var KM = document.myform.addmode[i];
				if (KM.checked==true)	   
					addmode = KM.value
				}
		
				if (addmode==0 && $('input[name=PicUrls]').val()=='')
				{
				  alert('���ϴ�<%=KS.C_S(ChannelID,3)%>��!');
				  $('input[name=imgurl1]').focus();
				  return false;
				}
				  $('form[name=myform]').submit();
				  $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
			}
		</script>
    
<%
			 .Write "</body>"
			 .Write "</html>"
			 End With
		End Sub
		
		'����
		Sub PictureSave()
		   Dim MoreStart,MoreEnd,MorePicUrl,MorePhotoUrl,I,SelectInfoList,HasInRelativeID
		  With Response
			
			Title = KS.G("Title")
			PictureContent= KS.FilterIllegalChar(Request.Form("Content")) : If KS.IsNul(PictureContent) Then PictureContent=" "
			Hits        = KS.ChkClng(KS.G("Hits"))
			HitsByDay   = KS.ChkClng(KS.G("HitsByDay"))
			HitsByWeek  = KS.ChkClng(KS.G("HitsByWeek"))
			HitsByMonth = KS.ChkClng(KS.G("HitsByMonth"))
			
			PhotoUrl     = KS.G("PhotoUrl")
			If KS.G("AddMode")="0" Then
			   PicUrls     = KS.G("PicUrls")
			Else
			   MoreStart=KS.ChkClng(KS.G("MoreStart"))
			   MoreEnd=KS.ChkClng(KS.G("MoreEnd"))
			   If MoreStart>MoreEnd Then .Write "<script>alert('������ӵĽ���ID�����С��ʼID!');history.back();</script>":.end
			   MorePicUrl=KS.G("MorePicUrl")
			   For I=MoreStart to MoreEnd
			    If PicUrls="" Then
				 PicUrls=Replace(MorePicUrl,"#",I)
				Else
				 PicUrls=PicUrls & "|||" & Replace(MorePicUrl,"#",I)
				End If
			   Next
			End If
			
			Recommend   = KS.ChkClng(KS.G("Recommend"))
			Rolls       = KS.ChkClng(KS.G("Rolls"))
			Strip       = KS.ChkClng(KS.G("Strip"))
			Popular     = KS.ChkClng(KS.G("Popular"))
			Comment     = KS.ChkClng(KS.G("Comment"))
			IsTop       = KS.ChkClng(KS.G("IsTop"))
			Slide       = KS.ChkClng(KS.G("Slide"))
			Makehtml    = KS.ChkClng(KS.G("Makehtml"))
			Tid = KS.G("Tid")
			SpecialID = Replace(KS.G("SpecialID")," ",""):SpecialID = Split(SpecialID,",")
			SelectInfoList = Replace(KS.G("SelectInfoList")," ","")
			Verific=1
			KeyWords = KS.G("KeyWords")
			Author  = KS.G("Author")
			Origin  = KS.G("Origin")
			AddDate = KS.G("AddDate")
			If Not IsDate(AddDate) Then AddDate=Now
			Rank = Trim(KS.G("Rank"))
			ShowStyle=KS.ChkClng(KS.G("ShowStyle"))
			PageNum=KS.ChkClng(KS.G("PageNum"))
				
				'�շ�ѡ��
				ReadPoint   = KS.ChkClng(KS.G("ReadPoint"))
				ChargeType  = KS.ChkClng(KS.G("ChargeType"))
				PitchTime   = KS.ChkClng(KS.G("PitchTime"))
				ReadTimes   = KS.ChkClng(KS.G("ReadTimes"))
				InfoPurview = KS.ChkClng(KS.G("InfoPurview"))
				arrGroupID  = KS.G("GroupID")
				DividePercent=KS.G("DividePercent"):IF Not IsNumeric(DividePercent) Then DividePercent=0
				
				TemplateID = KS.G("TemplateID")
				WapTemplateID=KS.G("WapTemplateID")
				Dim filetype:filetype=KS.ChkClng(KS.G("filetype"))
				Dim FnameType
				Dim RS_C:Set RS_C=Server.CreateObject("Adodb.RecordSet")
					RS_C.Open "Select top 1 * From KS_Class Where ID='" & Tid & "'",conn,1,1
					If Not RS_C.Eof Then
					    FnameType=RS_C("FnameType")
						If KS.ChkClng(KS.G("TemplateFlag"))=2 Or TemplateID="" Then TemplateID=RS_C("TemplateID"):WapTemplateID=RS_C("WapTemplateID")
						If FileType=0 Then
						  If Action = "Add" OR Action="Verify" Then
						   Fname=KS.GetFileName(RS_C("FsoType"), Now, "") & FnameType
						   End If
						End If
					End If
				RS_C.Close:Set RS_C=Nothing
				If filetype=1 Then Fname=KS.G("FileName")

    			Call KSCls.CheckDiyField(ChannelID,UserDefineFieldArr,ErrMsg)  '����Զ����ֶ�		
			 
			If Title = "" Then .Write ("<script>alert('ͼƬ���Ʋ���Ϊ��!');history.back(-1);</script>")
			If PhotoUrl = "" And KS.ChkClng(KS.S("autothumb"))=0 Then .Write ("<script>alert('ͼƬ����ͼ����Ϊ��!');history.back(-1);</script>")
			
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Tid = "" Then ErrMsg = ErrMsg & "[ͼƬ���]��ѡ! \n"
			If Title = "" Then ErrMsg = ErrMsg & "[ͼƬ����]����Ϊ��! \n"
			If Title <> "" And Tid <> "" And Action = "Add" Then
			  SqlStr = "select top 1 * from " & KS.C_S(ChannelID,2) & " where Title='" & Title & "' And Tid='" & Tid & "'"
			   RS.Open SqlStr, conn, 1, 1
				If Not RS.EOF Then
				 ErrMsg = ErrMsg & "������Ѵ��ڴ�ƪͼƬ! \n"
			   End If
			   RS.Close
			End If
			If ErrMsg <> "" Then
			   .Write ("<script>alert('" & ErrMsg & "');history.back(-1);</script>")
			   .End
			Else
			      If KS.ChkClng(KS.G("TagsTF"))=1 Then Call KSCls.AddKeyTags(KeyWords)
				  
			      If KS.ChkClng(KS.G("BeyondSavePic"))=1 Then
				  	SaveFilePath = KS.GetUpFilesDir & "/"
					KS.CreateListFolder (SaveFilePath)
				   Dim sPicUrlArr:sPicUrlArr=Split(PicUrls,"|||")
				   Dim sTemp,Url1,thumburl,ThumbFileName
				   PicUrls=""
				   For I=0 To Ubound(sPicUrlArr)
				     If Left(Lcase(Split(sPicUrlArr(i),"|")(1)),4)="http" and instr(Lcase(Split(sPicUrlArr(i),"|")(1)),lcase(ks.setting(2)))=0 Then
					    Dim PicURl:PicUrl=SaveFilePath & year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & i &".jpg"
						Url1=PicURL
					    Call KS.SaveBeyondFile(PicURL, Split(sPicUrlArr(i),"|")(1))
					    thumburl=replace(url1,ks.setting(2),"")
					    ThumbFileName=split(thumburl,".")(0)&"_S."&split(thumburl,".")(1)
						if instr(Lcase(thumburl),"http://")=0 Then
							Dim T:Set T=New Thumb
							Dim CreateTF:CreateTF=T.CreateThumbs(thumburl,ThumbFileName)
							if CreateTF=false Then
								ThumbFileName=url1
							end if
							Set T=Nothing
						end if
					  sTemp=Split(sPicUrlArr(i),"|")(0) & "|" & Url1 &"|" &ThumbFileName
					 Else
					  sTemp=sPicUrlArr(I)
					 End If
					 If I=0 Then
					   PicUrls=sTemp
					 Else
					   PicUrls=PicUrls & "|||" & sTemp
					 End If
				   Next
				   PhotoUrl= KS.ReplaceBeyondUrl(PhotoUrl, SaveFilePath)
				  End If
				  
				  If KS.ChkClng(KS.S("autothumb"))=1 And KS.IsNul(PhotoUrl) Then  PhotoUrl=Split(Split(PicUrls,"|||")(0),"|")(2)
				  
				  
				  If Action = "Add" Then
					SqlStr = "select top 1 * from " & KS.C_S(ChannelID,2) & " where 1=0"
					RS.Open SqlStr, conn, 1, 3
					RS.AddNew
					RS("Title") = Title
					RS("PhotoUrl") = PhotoUrl
					RS("PictureContent") = PictureContent
					RS("PicUrls") = PicUrls
					RS("Recommend") = Recommend
					RS("Rolls") = Rolls
					RS("Strip") = Strip
					RS("Popular") = Popular
					RS("Verific") = Verific
					RS("Comment") = Comment
					RS("IsTop") = IsTop
					RS("Tid") = Tid
					RS("KeyWords") = KeyWords
					RS("Author") = Author
					RS("Origin") = Origin
					RS("AddDate") = AddDate
					RS("Rank") = Rank
					RS("Slide") = Slide
					RS("TemplateID") = TemplateID
					RS("WapTemplateID")  = WapTemplateID
					RS("Hits") = Hits
					RS("HitsByDay") = HitsByDay
					RS("HitsByWeek") = HitsByWeek
					RS("HitsByMonth") = HitsByMonth
					RS("Fname") = Fname
					RS("Inputer") = KS.C("AdminName")
					RS("RefreshTF") = Makehtml
					RS("Score") = KS.ChkClng(KS.G("Score"))
					RS("DelTF") = 0
					RS("ShowStyle")=ShowStyle
					RS("PageNum")=PageNum
					RS("ReadPoint")=ReadPoint
				    RS("ChargeType")=ChargeType
				    RS("PitchTime")=PitchTime
				    RS("ReadTimes")=ReadTimes
					RS("InfoPurview")=InfoPurview
					RS("arrGroupID")=arrGroupID
					RS("DividePercent")=DividePercent
					If F_B_Arr(18)="1" Then	 RS("MapMarker")=KS.G("MapMark")
					Call KSCls.AddDiyFieldValue(RS,UserDefineFieldArr)
					RS.Update
					
				   'д��Session,�����һƪͼƬ����
				   Session("KeyWords") = KeyWords
				   Session("Author") = Author
				   Session("Origin") = Origin
                   RS.MoveLast
				   If Left(Ucase(Fname),2)="ID" Then
					   RS("Fname") = RS("ID") & FnameType
					   RS.Update
					End If
					
					For I=0 To Ubound(SpecialID)
						Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
					Next
					
					Call KSCls.UpdateRelative(ChannelID,RS("ID"),SelectInfoList,0)
 					Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,Tid,PictureContent,KeyWords,PhotoUrl,AddDate,KS.C("AdminName"),Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,RS("Fname"))
	 				'�����ϴ��ļ�
					 Call KS.FileAssociation(ChannelID,RS("ID"),PicUrls & PhotoUrl & PictureContent,0)

			        Call RefreshHtml(1)
					RS.Close:Set RS = Nothing
					
				ElseIf Action = "Edit" Or Action="Verify"  Then
				PicID = KS.ChkCLng(Request("PicID"))
				SqlStr = "SELECT top 1 * FROM " & KS.C_S(ChannelID,2) & " Where ID=" & PicID
					RS.Open SqlStr, conn, 1, 3
					If RS.EOF And RS.BOF Then
					 .Write ("<script>alert('�������ݳ���!');history.back(-1);</script>")
					 .End
					End If
					RS("Title") = Title
					RS("PhotoUrl") = PhotoUrl
					RS("PictureContent") = PictureContent
					RS("PicUrls") = PicUrls
					RS("Recommend") = Recommend
					RS("Rolls") = Rolls
					RS("Strip") = Strip
					RS("Popular") = Popular
					RS("Comment") = Comment
					RS("IsTop") = IsTop
					RS("Tid") = Tid
					RS("KeyWords") = KeyWords
					RS("Author") = Author
					RS("Origin") = Origin
					RS("AddDate") = AddDate
					RS("Rank") = Rank
					RS("ShowStyle")=ShowStyle
					RS("PageNum")=PageNum
					RS("Slide") = Slide
					RS("TemplateID") = TemplateID
					RS("WapTemplateID")  = WapTemplateID
					If Makehtml = 1 Then
					 RS("RefreshTF") = 1
					End If
					RS("Hits") = Hits
					RS("HitsByDay") = HitsByDay
					RS("HitsByWeek") = HitsByWeek
					RS("HitsByMonth") = HitsByMonth
					RS("Score") = KS.ChkClng(KS.G("Score"))
					RS("ReadPoint")=	ReadPoint
				    RS("ChargeType")=ChargeType
				    RS("PitchTime")=PitchTime
				    RS("ReadTimes")=ReadTimes
					RS("InfoPurview")=InfoPurview
					RS("arrGroupID")=arrGroupID
					RS("DividePercent")=DividePercent
					If Action="Verify" Then
					  Inputer=RS("Inputer")
					End If
					RS("Verific") = Verific
					If F_B_Arr(18)="1" Then	 RS("MapMarker")=KS.G("MapMark")
					Call KSCls.AddDiyFieldValue(RS,UserDefineFieldArr)
					RS.Update
                   RS.MoveLast
			       If TID<>Request.Form("OldClassID") Then
					     Call KSCls.DelInfoFile(ChannelID,Request.Form("OldClassID"), Split(RS("PicUrls"), "|||"),RS("Fname"))
				   End If
						Conn.Execute("Delete From KS_SpecialR Where InfoID=" & RS("ID") & " and channelid=" & ChannelID)
						For I=0 To Ubound(SpecialID)
						Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
						Next
					Call KSCls.UpdateRelative(ChannelID,PicID,SelectInfoList,1)
					Call LFCls.UpdateItemInfo(ChannelID,PicID,Title,Tid,PictureContent,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
	 				'�����ϴ��ļ�
					Call KS.FileAssociation(ChannelID,PicID,PicUrls & PhotoUrl & PictureContent,1)
				    Call RefreshHtml(2)
		          
				  RS.Close:Set RS = Nothing
					IF Action="Verify" Then     '��������Ͷ��ͼƬ�����û������мӻ��ֵȣ�������ǩ��ͼƬ����
							  '���û�������ֵ��������֪ͨ����
							  IF Inputer<>"" And Inputer<>KS.C("AdminName") Then Call KS.SignUserInfoOK(ChannelID,Inputer,Title,PicID)
							 .Write ("<script> parent.frames['MainFrame'].focus();alert('" & KS.C_S(ChannelID,3) &"�ɹ�ǩ��,ϵͳ�ѷ���һ��վ��֪ͨ�Ÿ�Ͷ����!');location.href='KS.ItemInfo.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) &" >> <font color=red>ǩ�ջ�Ա" & KS.C_S(ChannelID,3)) &"</font>';</script>") 
							 
				    End If
					If KeyWord <>"" Then
						 .Write ("<script> parent.frames['MainFrame'].focus();alert('" & KS.C_S(ChannelID,3) &"�޸ĳɹ�!');location.href='KS.Picture.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=PictureSearch&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) &" >> <font color=red>�������</font>") & "';</script>")
					End If
				End If
			End If
		 End With		
		End Sub
		
			Sub RefreshHtml(Flag)
			     Dim TempStr,EditStr,AddStr
			    If Flag=1 Then
				  TempStr="���":EditStr="�޸�" & KS.C_S(ChannelID,3):AddStr="�������" & KS.C_S(ChannelID,3)
				Else
				  TempStr="�޸�":EditStr="�����޸�" & KS.C_S(ChannelID,3):AddStr="���" & KS.C_S(ChannelID,3)
				End If
			    With Response
				     .Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
					 .Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>"
					 .Write " <Br><br><br><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
					  .Write "	  <tr class=""sort""> "
					  .Write "		<td  height=""28"" colspan=2>ϵͳ������ʾ��Ϣ</td>" & vbcrlf
					  .Write "	  </tr>"
                      .Write "    <tr class='tdbg'>"
					  .Write "          <td align='center'><img src='images/succeed.gif'></td>"
					  .Write "<td><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ϲ��" & TempStr &"" & KS.C_S(ChannelID,3) & "�ɹ���</b><br>"

					   If Makehtml = 1 Then
					      .Write "<div style=""margin-top:15px;border: #E7E7E7;height:220; overflow: auto; width:100%"">" 
					    If KS.C_S(ChannelID,7)=1 Or KS.C_S(ChannelID,7)=2 Then
						  	 .Write "<div><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						  Else
						  .Write "<div style=""height:25px""><li>����" & KS.C_S(ChannelID,1) & "û����������HTML�Ĺ��ܣ�����ID��Ϊ <font color=red>" & RS("ID") & "</font>  ��" & KS.C_S(ChannelID,3) & "û������!</li></div> "
						  End If
						  
							If KS.C_S(ChannelID,7)<>1 Then
							  .Write "<div style=""height:25px""><li>����" & KS.C_S(ChannelID,1) & "����Ŀҳû����������HTML�Ĺ��ܣ�����ID��Ϊ <font color=red>" & TID & "</font>  ����Ŀû������!</li></div> "
							Else
							 If KS.C_S(ChannelID,9)<>1 Then
								  Dim FolderIDArr:FolderIDArr=Split(left(KS.C_C(Tid,8),Len(KS.C_C(Tid,8))-1),",")
								  For I=0 To Ubound(FolderIDArr)
								  .Write "<div align=center><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true'></iframe></div>"
								   Next
							 End If
						   End If
					   If Split(KS.Setting(5),".")(1)="asp" or KS.C_S(ChannelID,9)<>3 Then
					   ' .Write "<div style=""margin-left:140;color:blue;height:25px""><li>���� <a href=""" & KS.GetDomain & """ target=""_blank""><font color=red>��վ��ҳ</font></a> û����������HTML�Ĺ��ܻ򷢲�ѡ��û�п���������û������!</li></div>"
					   Else
					     .Write "<div align=center><iframe src=""Include/RefreshIndex.asp?RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
					   End If
					   .Write "</div>"
					 End If
					  .Write   "</td></tr>"
					  .Write "	  <tr class='tdbg'>"
					  .Write "		<td height=""25"" colspan=""2"" style=""text-align:right"">��<a href=""#"" onclick=""location.href='KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&Action=Edit&KeyWord=" & KeyWord &"&SearchType=" & SearchType &"&StartDate=" & StartDate & "&EndDate=" & EndDate &"&ID=" & RS("ID") & "';""><strong>" & EditStr &"</strong></a>��&nbsp;��<a href=""#"" onclick=""location.href='KS.Picture.asp?ChannelID=" & ChannelID & "&Action=Add&FolderID=" & Tid & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr="&server.URLEncode("���" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & Tid & "';""><strong>" & AddStr & "</strong></a>��&nbsp;��<a href=""#"" onclick=""location.href='KS.ItemInfo.asp?ID=" & Tid & "&ChannelID=" & ChannelID & "&Page=" & Page&"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & Tid & "';""><strong>" & KS.C_S(ChannelID,3) & "����</strong></a>��&nbsp;��<a href=""" & KS.GetDomain &"Item/Show.asp?M=" & ChannelID & "&D=" & RS("ID") & """ target=""_blank""><strong>Ԥ��" & KS.C_S(ChannelID,3) & "����</strong></a>��</td>"
					  .Write "	  </tr>"
					  .Write "	</table>"				
			End With
		End Sub

End Class
%> 
