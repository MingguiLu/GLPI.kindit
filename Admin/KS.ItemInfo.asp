<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
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
Set KSCls = New Admin_ItemInfo
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ItemInfo
        Private KS,ComeUrl,KSCls
		'=====================================������ҳ��ȫ�ֱ���==============================================================
        Private ID, I, totalPut, Page, RS,ComeFrom,ItemManageUrl
		Private KeyWord, SearchType, StartDate, EndDate,SearchParam, MaxPerPage,T, TitleStr, VerificStr
		Private TypeStr, AttributeStr, FolderID, TemplateID,FolderName, Action,TotalPages
		Private FileName,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj,SaveFilePath
		Private ChannelID,F_B_Arr,F_V_Arr,IXML,INode,O
		Private XmlFields,XmlFieldArr,FI
		'======================================================================================================================

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
		If ChannelID=0 Then ChannelID=1
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
		F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
		
		Select Case KS.C_S(ChannelID,6)
		 Case 2:ItemManageUrl="KS.Picture.asp"
		 Case 3:ItemManageUrl="KS.Down.asp"
		 Case 4:ItemManageUrl="KS.Flash.asp"
		 Case 5:ItemManageUrl="KS.Shop.asp"
		 Case 7:ItemManageUrl="KS.Movie.asp"
		 Case 8:ItemManageUrl="KS.Supply.asp"
		 Case Else:ItemManageUrl="KS.Article.asp"
		End Select
		
		KeyWord    = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		StartDate  = KS.G("StartDate")
		EndDate    = KS.G("EndDate")
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
		Page = KS.ChkClng(KS.G("page")) : If Page=0 Then  Page = 1
		O = KS.ChkClng(KS.G("O"))
		
		Select Case Action
		 Case "Recely"
           If Not KS.ReturnPowerResult(0, "M010006") Then 
		    Call KS.ReturnErr(1, "")
		   Else
             Call KSCls.Recely(ChannelID)
           End If
		 Case "RecelyBack"
		    Call KSCls.RecelyBack(ChannelID)
		 Case "Delete"
			If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10004") Then  
			 Call KS.ReturnErr(1, "")
			Else
		    Call KSCls.DelBySelect(ChannelID)
			End If
		 Case "DeleteAll"
			If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10004") Then  
			 Call KS.ReturnErr(1, "")
			Else
		    Call KSCls.DeleteAll() 
			End If
		 Case "VerifyAll"
            Call KSCls.VerificAll(ChannelID)
		 Case "Tuigao"
		    Call KSCls.Tuigao(ChannelID)
		 Case "BatchSet"
		    Call KSCls.BatchSet(ChannelID)
		 Case "JS"
		   If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10007") Then  
			  Call KS.ReturnErr(0, "")
			Else
			  Call AddToJS()
			End If
		 Case "Special"
		  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10006") Then 
			 Call KS.ReturnErr(0, "")
		  Else
		     Call KSCls.AddToSpecial(ChannelID)
		  End If
		 Case "SetAttribute"
			If Not KS.ReturnPowerResult(0, "M010005") Then 
				 Call KS.ReturnErr(1, "")
			Else
		         Call SetAttribute()
			End If
		 Case "MoveClass"
		    If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then
				 Call KS.ReturnErr(1, "")
			Else
		         Call MoveClass()
			End If
		 Case "Paste"
		  	If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10011") Then  
			   Call KS.ReturnErr(1, "")   
            Else
		       Call KSCls.Paste(ChannelID)
			End If 
		 Case Else
		       Call ItemInfoMain()
		End Select
		
	 End Sub
	 
	 Sub ItemInfoMain()
		ID = KS.G("ID"):If ID = "" Then ID = "0"
		MaxPerPage = Cint(KS.C_S(ChannelID,11))     'ȡ��ÿҳ��ʾ����
		With KS
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<title>������ҳ��</title>"
		.echo "<link href=""include/admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		.echo "<script language=""JavaScript"">"
		.echo " var ClassID='" & ID & "';                //Ŀ¼ID" & vbCrLf
		.echo " var Page='" & Page & "';                 //��ǰҳ��" & vbCrLf
		.echo " var KeyWord='" & KeyWord & "';           //�ؼ���" & vbCrLf
		.echo " var SearchParam='" & SearchParam & "';   //������������" & vbCrLf
		
		.echo "</script>" & vbCrLf
		.echo "<script language=""JavaScript"" src=""../KS_Inc/Common.js""></script>" & vbCrLf
		.echo "<script language=""JavaScript"" src=""../KS_Inc/JQuery.js""></script>" & vbCrLf
		.echo "<script language=""JavaScript"" src=""../KS_Inc/kesion.box.js""></script>" & vbCrLf
		%>
		<script language="JavaScript">
		function ClassToggle(f)
		{
		  setCookie("classExtStatus",f)
		  $('#classNav').toggle('slow');
		  $('#classOpen').toggle('show');
		}
		function ProcessTuigao(ev,Id)
		{
		    var ids=get_Ids(document.myform);
			if (Id=='') Id=ids;
			if (Id=='')
			{
			  alert('�Բ�����û��ѡ��Ҫ�˸���ĵ�!');
			  return;
			}
			var p=new KesionPopup();
		 	p.mousepop("<b>�˸�ԭ��</b>","<div style='height:200px;text-align:center'><form name='rform' action='KS.ItemInfo.asp?action=Tuigao&Page=<%=Page%>' method='post'><input type='hidden' name='channelid' value='<%=ChannelID%>'><input type='hidden' name='Id' value='"+Id+"'><textarea name='AnnounceContent' style='width:300px;height:130px'>����{$UserName}���������ĸ����{$Title}�������ϱ�վҪ�����޸ĺ��������ύ��</textarea><br><br/><label><input type='checkbox' value='1' name='Email' checked>����վ�ڶ���֪ͨ</label> <input type='submit' value='ȷ���˸�' class='button'> <input type='submit' value='ȡ���˸�' class='button' onclick='closeWindow();'></form></div>",400);

		}
		function CreateHtml(){  
		   var ids=get_Ids(document.myform);
			if (ids!='')
		new KesionPopup().PopupCenterIframe('����ѡ���ĵ�','Include/RefreshHtmlSave.Asp?ChannelID=<%=ChannelID%>&Types=Content&RefreshFlag=IDS&ID='+ids,530,110,'no')
			else 
			alert('��ѡ��Ҫ�������ĵ�!');
		}	
		function MoveClass(){
		   var ids=get_Ids(document.myform);
			if (ids!='')
		new KesionPopup().PopupCenterIframe('�����ƶ�ѡ���ĵ�','KS.ItemInfo.asp?ChannelID=<%=ChannelID%>&action=MoveClass&ID='+ids,530,110,'no')
			else 
			alert('��ѡ��Ҫ�ƶ����ĵ�!');
		}	
		function CreateNews(){   
		   location.href='<%=ItemManageUrl%>?ChannelID=<%=ChannelID%>&Action=Add&FolderID='+ClassID;
           $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=<%=ChannelID%>&OpStr='+escape("���<%=KS.C_S(ChannelID,3)%>")+'&ButtonSymbol=AddInfo&FolderID='+ClassID;
		}
		function VerifyInfo()
		{
		   location.href='KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=<%=ChannelID%>';
           $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=<%=ChannelID%>&OpStr='+escape("ǩ��<%=KS.C_S(ChannelID,3)%>")+'&ButtonSymbol=Disabled';
		}
		function Edit()
		{   var ids=get_Ids(document.myform);
			 if (ids!='')
					 if (ids.indexOf(',')==-1){
						 location.href='<%=ItemManageUrl%>?Page='+Page+'&Action=Edit&'+SearchParam+'&ID='+ids;
						 if (KeyWord=='')
							$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("�༭<%=KS.C_S(ChannelID,3)%>")+'&ButtonSymbol=AddInfo&FolderID='+ClassID;
						 else
						   $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("<%=KS.C_S(ChannelID,1)%> >> ������� >> <font color=red>�༭<%=KS.C_S(ChannelID,3)%></font>")+'&ButtonSymbol=AddInfo';
						 }
					   else alert('һ��ֻ�ܹ��༭һ<%=KS.C_S(ChannelID,4)%><%=KS.C_S(ChannelID,3)%>');
					 
			else 
			{
			alert('��ѡ��Ҫ�༭��<%=KS.C_S(ChannelID,3)%>');
			}
		}
		function Recely()
		{ 
		   if (chk_idBatch(myform,'��ѡ�е�<%=KS.C_S(ChannelID,3)%>�������վ��')==true)
		   {
		    $('input[name=action]').val('Recely'); 
			$('form[name=myform]').submit();
		   }
		}
		function BackRecely()
		{
		   if (chk_idBatch(myform,'��ѡ�е�<%=KS.C_S(ChannelID,3)%>��ԭ��')==true)
		   {
		    $('input[name=action]').val('RecelyBack'); 
			$('form[name=myform]').submit();
		   }
		}
		function Delete()
		{ 
		   if (chk_idBatch(myform,'�˲���������,����ɾ��ѡ�е�<%=KS.C_S(ChannelID,3)%>��')==true)
		   {
		    $('input[name=action]').val('Delete'); 
			$('form[name=myform]').submit();
		   }
		}
		function DelAll()
		{
		  if (confirm('������ʾ:\n\nһ����ս��������ģ����Ļ���վ�ĵ�,�Ҵ˲��������棬ȷ����ջ���վ��')==true)
		  {
		    $('input[name=action]').val('DeleteAll');
			$('form[name=myform]').submit();
		  }
		}
		function VerificAll()
		{
		   if (chk_idBatch(myform,'ȷ���������ѡ�е�<%=KS.C_S(ChannelID,3)%>��')==true)
		   {
		    $('input[name=action]').val('VerifyAll'); 
			$('form[name=myform]').submit();
		   }

		}
		function Tuigao()
		{
		  ProcessTuigao(event,'')
		  return;
		
		 if (chk_idBatch(myform,'ȷ�������˸�ѡ�е�<%=KS.C_S(ChannelID,3)%>��')==true)
		   {
		    $('input[name=action]').val('Tuigao'); 
			$('input[name=myform]').submit();
		   }
		}
		
		function Copy()
		{
		    var ids=get_Ids(document.myform);
			if (ids!='')
			  {
			   top.CommonCopyCut.ChannelID=<%=ChannelID%>;
			   top.CommonCopyCut.PasteTypeID=2;
			   top.CommonCopyCut.SourceFolderID=ClassID;
			   top.CommonCopyCut.FolderID='0';
			   top.CommonCopyCut.ContentID=ids;
			  }
			else
			 alert('��ѡ��Ҫ���Ƶ�<%=KS.C_S(ChannelID,3)%>!');
		}
		function Paste()
		{ 
		  if (ClassID=='0')
		   { top.CommonCopyCut.PasteTypeID=0;
			 alert('Ŀ��Ŀ¼������!');
			}
		  if (top.CommonCopyCut.ChannelID==<%=ChannelID%> && top.CommonCopyCut.PasteTypeID!=0)
		   {  var Param='';
			  Param='?ChannelID=<%=ChannelID%>&Action=Paste&Page='+Page;
			  Param+='&PasteTypeID='+top.CommonCopyCut.PasteTypeID+'&DestFolderID='+ClassID+'&SourceFolderID='+top.CommonCopyCut.SourceFolderID+'&FolderID='+top.CommonCopyCut.FolderID+'&ContentID='+top.CommonCopyCut.ContentID;
			  if (top.CommonCopyCut.PasteTypeID==2) //����
			 {
				location.href='KS.ItemInfo.asp'+Param;
			 }
			else
			 alert('�Ƿ�����!');
		   }
		  else
		   alert('ϵͳ���а�û������!');
		}
		function AddToSpecial()
		{  var ids=get_Ids(document.myform);
			if (ids!='')
				{     
				OpenWindow('KS.Frame.asp?PageTitle='+escape('<%=KS.C_S(ChannelID,3)%>���뵽ר��')+'&URL=KS.ItemInfo.asp&ChannelID=<%=ChannelID%>&Action=Special&NewsID='+ids,300,350,window);
				}
			else alert('��ѡ��Ҫ����ר���<%=KS.C_S(ChannelID,3)%>!');
			Select(2);
		}
		function AddToJS()
		{  var ids=get_Ids(document.myform);
			if (ids!='')
				{     
				OpenWindow('KS.Frame.asp?PageTitle='+escape('<%=KS.C_S(ChannelID,3)%>���뵽����JS')+'&URL=KS.ItemInfo.asp&ChannelID=<%=ChannelID%>&Action=JS&NewsID='+ids,300,100,window);
				}
			else alert('��ѡ��Ҫ����JS��<%=KS.C_S(ChannelID,3)%>!');
			Select(2);
		}
		function SetAttribute()
		{   var ids=get_Ids(document.myform);
		     if (ids=='')
			 {
			  alert('��ѡ��Ҫ�������Ե�<%=KS.C_S(ChannelID,3)%>!');
			  return;
			 }
			 OpenWindow('KS.Frame.asp?PageTitle='+escape('��������<%=KS.C_S(ChannelID,3)%>����')+'&URL=KS.ItemInfo.asp&ChannelID=<%=ChannelID%>&Action=SetAttribute&ID='+ids,500,420,window);
			 window.location.reload();
		}
		function MoveToClass()
		{   var ids=get_Ids(document.myform);
		     if (ids=='')
			 {
			  alert('��ѡ��Ҫ�����ƶ���<%=KS.C_S(ChannelID,3)%>!');
			  return;
			 }
			 OpenWindow('KS.Frame.asp?PageTitle='+escape('<%=KS.C_S(ChannelID,3)%>�����ƶ�<%=KS.C_S(ChannelID,3)%>')+'&URL=KS.Class.asp&ChannelID=<%=ChannelID%>&Action=MoveInfo&From=main&ID='+ids,500,400,window);
			 window.location.reload();		
		}
		function ViewArticle(ArticleID)
		{
		window.open ('../Item/Show.asp?m=<%=ChannelID%>&d='+ArticleID);
		}
		function setstatus(Obj)
		  {var today=new Date()
			if (Obj.nextSibling.style.display=='none')
			 {
			  Obj.nextSibling.style.display='';
			  $('#StartDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-01');
			  $('#EndDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate());
			 }
			else 
			{
			 Obj.nextSibling.style.display='none';
			 $('#StartDate').val('');
			 $('#EndDate').val('');
			 }
		}
		function set(o,v)
		{
		 
		 if (parseInt(v)!=0)
		  {
		  var ids=get_Ids(document.myform);
		  if (ids!='')
		   {
					if (confirm('ȷ����ѡ�е�<%=KS.C_S(ChannelID,3)%>'+o.value)==true)
					{
					    $('#SetAttributeBit').val(v);
						$('input[name=action]').val('BatchSet'); 
						$('form[name=myform]').submit();

					}
			}
		   else
		    alert('��ѡ��Ҫ���õ�<%=KS.C_S(ChannelID,3)%>');
		  }
		}
		function GetKeyDown()
		{
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {      case 90 : Select(2); break;
			 case 77 : CreateNews();break;
			 case 65 : Select(0);break;
			 case 83 : AddToSpecial();break;
			 case 74 : AddToJS();break;
			 case 85 : SetAttribute();break;
			 case 67 : 
				{event.keyCode=0;event.returnValue=false;Copy();}
                 break;
			 case 86 : 
			   if (top.CommonCopyCut.ChannelID==<%=ChannelID%> && top.CommonCopyCut.PasteTypeID!=0 && ClassID!='0')
			   { event.keyCode=0;event.returnValue=false;Paste();}
			   else
			    {
				 if (top.CommonCopyCut.PasteTypeID!=0)
				alert('��ת��Ŀ����Ŀ����ճ��!');
				return;
				}
				break;
			 case 69 : event.keyCode=0;event.returnValue=false;Edit();break;
			 case 68 : Recely();break;
			 case 70 : event.keyCode=0;event.returnValue=false;parent.frames['LeftFrame'].initializeSearch('<%=KS.C_S(ChannelID,1)%>',<%=ChannelID%>,<%=KS.C_S(ChannelID,6)%>)
		   }	
		else if (event.keyCode==46) Delete();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body scroll=no onkeydown=""GetKeyDown();"" onselectstart=""return false;"" style='background:#fff;'>"
		.echo "<ul id='menu_top'>"
		If ComeFrom="RecycleBin" Then
		 .echo "<li class='parent' onclick='BackRecely()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/reb.gif' border='0' align='absmiddle'>������ԭ</span></li>"
		 .echo "<li class='parent' onclick='Delete()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>����ɾ��</span></li>"
		 .echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" onclick='DelAll()'><img src='images/ico/recyclebin.gif' border='0' align='absmiddle'>һ����ջ���վ</span></li>"
		ElseIf ComeFrom="Verify" Then
		    If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10012") Then 
		    Call KS.ReturnErr(1, "")
			End If

		 .echo "<li class='parent' onclick='VerificAll()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�������</span></li>"
		 .echo "<li class='parent' onclick='Tuigao()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>�����˸�</span></li>"
		 .echo "<li class='parent' onclick='Recely()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/recycl.gif' border='0' align='absmiddle'>�������վ</span></li>"
		 .echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" onclick='Delete()'><img src='images/ico/del.gif' border='0' align='absmiddle'>����ɾ��</span></li>"
		Else
		.echo "<li class='parent' onclick='CreateNews();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>���" & KS.C_S(ChannelID,3) & "</span></li>"
		.echo "<li class='parent' onclick='VerifyInfo();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>ǩ��" & KS.C_S(ChannelID,3) & "</span></li>"
		.echo "<li class='parent' onclick='Recely()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/recycl.gif' border='0' align='absmiddle'>�������վ</span></li>"
		.echo "<li class='parent' onclick='Delete()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>����ɾ��</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""������������"" onclick=""SetAttribute();""><img src='images/ico/set.gif' border='0' align='absmiddle'>��������</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""�����ƶ�""  onClick=""MoveToClass();""><img src='images/ico/move.gif' border='0' align='absmiddle'>�����ƶ�</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""��������JS"" onclick=""AddToJS();""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>����JS</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""����ר��""  onClick=""AddToSpecial();""><img src='images/ico/as.gif' border='0' align='absmiddle'>����ר��</span></li>"
        End If
			.echo "<li></li><div><select OnChange=""location.href='KS.ItemInfo.asp?ComeFrom=" & ComeFrom & "&ChannelID=" & ChannelID & "&id='+this.value;$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID='+this.value;"" style='width:150px' name='id'>"
			.echo "<option value=''>������ת��...</option>"
			.echo Replace(KS.LoadClassOption(ChannelID,false),"value='" & ID & "'","value='" & ID &"' selected") & " </select>"

			
		   .echo "</div>"
		   .echo (" </ul>")
		
			.echo ("<div style=""*height:92%;overflow-x: hidden; overflow-y: scroll; width:100%"">")
		 If KeyWord<>"" or (StartDate <> "" And EndDate <> "") Then
		 .echo ("<img src='Images/ico/search.gif' align='absmiddle'> �������: ")
				 If StartDate <> "" And EndDate <> "" Then
					.echo (KS.C_S(ChannelID,3) & "���������� <font color=red>" & StartDate & "</font> �� <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
				 End If
				 If  KeyWord<>"" Then
				   Select Case SearchType
					Case 0:.echo ("�ĵ������к��� <font color=red>" & KeyWord & "</font> ��" & KS.C_S(ChannelID,3))
					Case 1:.echo ("�ĵ�¼��Ա�к��� <font color=red>" & KeyWord & "</font> ��" & KS.C_S(ChannelID,3))
					Case 2:.echo ("�ĵ��ؼ����к���<font color=red>" & KeyWord & "</font> ��" & KS.C_S(ChannelID,3))
					Case 3:.echo ("�ĵ����ߺ���<font color=red>" & KeyWord & "</font> ��" & KS.C_S(ChannelID,3))
					Case 4:.echo ("��Ʒ��ź���<font color=red>" & KeyWord & "</font> ��" & KS.C_S(ChannelID,3))
					Case 5:.echo ("����Ʒ�ƺ���<font color=red>" & KeyWord & "</font> ��" & KS.C_S(ChannelID,3))
				  End Select
			     End If
		End If
		
		 If .G("ComeFrom")="RecycleBin" Then
		  ShowChannelList 
		 Else
	      ShowClassList ChannelID,ID
		 End If
		
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td><div style='margin:5px'><b>�鿴��</b>")
		 .echo ("<a href='?ChannelID=" & ChannelID & "&ComeFrom=" & ComeFrom & "'><font color=#999999>ȫ��</font></a> - ")
		 If ComeFrom="Verify" Then
		 .echo ("<a href='?ChannelID=" & ChannelID & "&Verific=0&ComeFrom=" & ComeFrom & "'><font color=#999999>�����" & KS.C_S(ChannelID,3) & "</font></a> - <a href='?ChannelID=" & ChannelID & "&Verific=2&ComeFrom=" & ComeFrom & "'><font color=#999999>��Ա�ݸ��" & KS.C_S(ChannelID,3) & "</font></a> - <a href='?ChannelID=" & ChannelID & "&Verific=3&ComeFrom=" & ComeFrom & "'><font color=#999999>���˸��" & KS.C_S(ChannelID,3) & "</font></a></div></td><td align='right'>")
		 Else
		 .echo ("<a href='?ChannelID=" & ChannelID & "&status=1&ComeFrom=" & ComeFrom & "'><font color=#999999>�Ƽ�</font></a> - <a href='?ChannelID=" & ChannelID & "&status=2&ComeFrom=" & ComeFrom & "'><font color=#999999>�õ�</font></a> - <a href='?ChannelID=" & ChannelID & "&status=3&ComeFrom=" & ComeFrom & "'><font color=#999999>����</font></a> - <a href='?ChannelID=" & ChannelID & "&status=4&ComeFrom=" & ComeFrom & "'><font color=#999999>�̶�</font></a> - <a href='?ChannelID=" & ChannelID & "&status=5&ComeFrom=" & ComeFrom & "'><font color=#999999>����</font></a> - <a href='?ChannelID=" & ChannelID & "&status=6&ComeFrom=" & ComeFrom & "'><font color=#999999>ͷ��</font></a> - <a href='?ChannelID=" & ChannelID & "&status=7&ComeFrom=" & ComeFrom & "'><font color=#999999>����</font></a>")
		 If KS.C_S(ChannelID,6)=1 Then
		 .echo (" - <a href='?ChannelID=" & ChannelID & "&status=11&ComeFrom=" & ComeFrom & "'><font color=#999999>��Ƶ</font></a>")
		 .echo (" - <a href='?ChannelID=" & ChannelID & "&status=10&ComeFrom=" & ComeFrom & "'><font color=#999999>ǩ��</font></a>")
		 End If
		  If ChannelID=5 Then 
		   .echo " - <a href='?ChannelID=" & ChannelID & "&status=8&ComeFrom=" & ComeFrom & "'><font color=green>�Ǽۡ�</font></a>"
		   .echo " - <a href='?ChannelID=" & ChannelID & "&status=9&ComeFrom=" & ComeFrom & "'><font color=#ff3300>���ۡ�</font></a>"
		  End If
		  Dim OrderArray:OrderArray=array("Ĭ������id��|id|1","�ĵ�id��|id|0","�������|hits|1","�������|hits|0","����ʱ���|adddate|1","����ʱ���|adddate|0")
		  dim t:t=ubound(OrderArray)
		  If ChannelID=5 Then
		   redim preserve OrderArray(t+10)
		   OrderArray(t+1)="�������|TotalNum|1" : OrderArray(t+2)="�������|TotalNum|0"
		   OrderArray(t+3)="�г��ۡ�|Price_Market|1": OrderArray(t+4)="�г��ۡ�|Price_Market|0"
		   OrderArray(t+5)="��ǰ���ۼۡ�|Price|1": OrderArray(t+6)="��ǰ���ۼۡ�|Price|0"
		   OrderArray(t+7)="ԭʼ���ۼۡ�|Price_Original|1" :OrderArray(t+8)="ԭʼ���ۼۡ�|Price_Original|0"
		   OrderArray(t+9)="��Ա�ۡ�|Price_Member|1" :OrderArray(t+10)="��Ա�ۡ�|Price_Member|0"
		 ElseIf Cint(KS.C_S(ChannelID,6))=3 Then
		   redim preserve OrderArray(t+6)
		   OrderArray(t+1)="�����ء�|HitsByDay|1" : OrderArray(t+2)="�����ء�|HitsByDay|0" 
		   OrderArray(t+3)="�����ء�|HitsByWeek|1" : OrderArray(t+4)="�����ء�|HitsByWeek|0" 
		   OrderArray(t+5)="�����ء�|HitsByMonth|1" : OrderArray(t+6)="�����ء�|HitsByMonth|0" 
		 End If
		  
		 .echo " <b>����</b><select onchange=""location.href=this.value"">"
		  for i=0 to ubound(OrderArray)
		    dim orderarr:orderarr=split(OrderArray(i),"|")
			 if O=i then
			.echo "<option selected value='KS.ItemInfo.asp?channelid=" & channelid &"&id=" & id & "&comefrom=" & comefrom & "&o=" & i &"'>" & orderarr(0) & "</option>"
			 else
			.echo "<option value='KS.ItemInfo.asp?channelid=" & channelid &"&id=" & id & "&comefrom=" & comefrom & "&o=" & i &"'>" & orderarr(0) & "</option>"
			 end if
		  next
		 .echo "</select>"
		 .echo ("</div></td><td align='right'>")
		 End If
		 .echo("<b>" & KS.C_S(ChannelID,1) & "</b>[���� <font color=red>" & Conn.Execute("select count(id) from " & KS.C_S(ChannelID,2) & " where verific=1")(0) & "</font> " & KS.C_S(ChannelID,4) & " ����վ <font color=blue>" &Conn.Execute("select count(id) from " & KS.C_S(ChannelID,2) & " where verific=1 and deltf=1")(0)  &"</font> "& KS.C_S(ChannelID,4) & "]</td></tr></table>")
		 
		XmlFields=LFCls.GetConfigFromXML("managemodelfield","/modelfield/model",ChannelID)
        If Not KS.IsNul(XmlFields) Then
		 XmlFieldArr=Split(XmlFields,",")
		End If
		.echo ("<div style=""width:100%;padding-bottom:15px;overflow-x: auto; height:auto"">")
		 .echo ("<table width=""100%"" align='center' border=""0"" cellpadding=""0"" cellspacing=""0"">")
		 .echo ("<form name='myform' method='Post' action='?channelid="& channelid & "'>")
		 .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
		 .echo ("<input type='hidden' name='SetAttributeBit' id='SetAttributeBit' value='0'>")
		 .echo ("<tr align=""center"" class=""sort"">")
		 .echo ("<td width='35' align='center' nowrap>ѡ��</td>")
			 If ChannelID=8 Then
			  .echo ("<td width='60'>����</td>")
			 End If
		.echo ("<td height=15 nowrap>����</td>")
		 If IsArray(XmlFieldArr) Then
			 For Fi=0 To Ubound(XmlFieldArr)
			   .echo ("<td nowrap>" & Split(XmlFieldArr(fi),"|")(0) & "</td>")
			 Next
			 if ComeFrom<>"" Then .Echo ("<td width='60' nowrap>״ ̬</td>")
        Else
			 .echo ("<td width=100>¼ ��</td><td width=80>�޸�����</td><td width=60> �� �� </td><td width=100> �� �� </td>")
			 If ComeFrom="" Then
			 .Echo ("<td width='60'>�� ��</td>")
			 Else
			 .Echo ("<td width='60'>״ ̬</td>")
			 End If
	   End If 
		 .Echo "<td> �� �� </td></tr>"

		   Dim Param
		   If ComeFrom="RecycleBin" Then
		    Param = Param & " DelTF=1"
		   ElseIf ComeFrom="Verify" Then
		    Param = Param & " DelTF=0 And Verific=" & KS.ChkClng(KS.G("Verific"))
		   Else
		    Param = Param & " DelTF=0  And Verific=1"
		   End If
		   
		   '�ǳ�������Ա��ֻ�ܹ����Լ���ӵ���Ϣ
		   'If KS.C("SuperTF")<>"1" Then	 Param=Param & " and inputer='" & KS.C("AdminName") & "'"
		   
		    If KS.C("SuperTF")<>"1" and Instr(KS.C("ModelPower"),KS.C_S(ChannelID,10)&"1")=0 Then 
			 If DataBaseType=1 Then
				 Param=Param & " and tid in(select id from ks_class where ','+cast(AdminPurview as nvarchar(4000))+',' like '%," & KS.C("AdminName") & "%'"
				 if (ID<>"0") then Param = Param & " And Ts Like '%" & ID & "%'" 
				 Param=Param & ")"
			 Else
				 Param=Param & " and tid in(select id from ks_class where ','+AdminPurview+',' like '%," & KS.C("AdminName") & "%'"
				 if (ID<>"0") then Param = Param & " And Ts Like '%" & ID & "%'" 
				 Param=Param & ")"
			 End If
			Elseif (ID<>"0") then 
			  Param = Param & " And Tid In (" & KS.GetFolderTid(ID) & ")" 
			End If
		  

		   If KeyWord <> "" or (StartDate <> "" And EndDate <> "") Then
		        If KeyWord<>"" Then
				Select Case SearchType
				  Case 0:Param = Param & " And (Title like '%" & KeyWord & "%')"
				  Case 1:Param = Param & " And Inputer like '%" & KeyWord & "%'"
				  Case 2:Param = Param & " And KeyWords like '%" & KeyWord & "%'"
				  Case 3:Param = Param & " And Author like '%" & KeyWord & "%'"
				  Case 4:Param = Param & " And ProID Like '%" & KeyWord & "%'"
				  Case 5:Param = Param & " And BrandID in(select id From KS_ClassBrand Where BrandName Like '%" & KeyWord & "%' or BrandeName Like '%" & KeyWord & "%')"
				End Select
				End If
				If StartDate <> "" And EndDate <> "" Then
					If CInt(DataBaseType) = 1 Then         'Sql
					   Param = Param & " And (AddDate>= '" & StartDate & "' And AddDate<= '" & DateAdd("d", 1, EndDate) & "')"
					Else                                                 'Access
					   Param = Param & " And (AddDate>=#" & StartDate & "# And AddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
					End If
				End If
		  Else
			select case KS.ChkClng(KS.S("Status"))
			 case 1 Param = Param & " And Recommend=1"
			 case 2 Param = Param & " And Slide=1"
			 case 3 Param = Param & " And Popular=1"
			 case 4 Param = Param & " And IsTop=1"
			 case 5 Param = Param & " And Comment=1"
			 case 6 Param = Param & " And Strip=1"
			 case 7 Param = Param & " And Rolls=1"
			 case 8 Param = Param & " And ProductType=2"
			 case 9 Param = Param & " And ProductType=3"
			 case 10 Param = Param &" And IsSign=1"
			 case 11 Param = Param &" And IsVideo=1"
			end select
			
		  End If
		  
		 
		
		Dim FieldStr:FieldStr="ID,Tid,Title,Inputer,AddDate,PhotoUrl,Verific,Recommend,Popular,Strip,Rolls,Slide,IsTop,Hits"

		If ChannelID=5 Then
		 FieldStr=FieldStr & ",ProductType"
		ElseIf ChannelID=8 Then
		 FieldStr=FieldStr & ",TypeID"
		ElseIf KS.C_S(ChannelID,6)=1 Then
		 FieldStr=FieldStr & ",IsVideo"
		End If
		If KS.ChkClng(KS.S("Status"))=10 Then
		 FieldStr=FieldStr & ",SignUser"
		End If
		
		If IsArray(XmlFieldArr) Then
		 For Fi=0 To Ubound(XmlFieldArr)
		  if lcase(Split(XmlFieldArr(fi),"|")(1))<>"modeltype" and lcase(Split(XmlFieldArr(fi),"|")(1))<>"attribute" and ks.foundinarr(lcase(FieldStr),lcase(Split(XmlFieldArr(fi),"|")(1)),",")=false then
		   FieldStr=FieldStr & "," & Split(XmlFieldArr(fi),"|")(1)
		  end if
		 Next
        End If
		
		totalPut = Conn.Execute("Select count(id) from [" & KS.C_S(ChannelID,2) & "] where " & Param)(0)
		if (TotalPut mod MaxPerPage)=0 then
			TotalPages = TotalPut \ MaxPerPage
		else
			TotalPages = TotalPut \ MaxPerPage + 1
		end if
		if Page > TotalPages and TotalPages<>0 then Page=TotalPages
		Dim OrderField,OrderType
		If IsArray(OrderArray) Then
		if O<=ubound(OrderArray) Then
		  OrderField=Split(OrderArray(O),"|")(1)
		  OrderType=Split(OrderArray(O),"|")(2)
		Else
		  OrderField="id":OrderType=1
		End If
	   Else
	      OrderField="id":OrderType=1
	   End If
		If OrderField<>"id" Then   '����������
		    Dim AscDesc:If OrderType=1 Then AscDesc=" Desc" Else AscDesc=" Asc"
			SQLStr="Select " & FieldStr & " From " & KS.C_S(ChannelID,2) & " where " & Param & " Order By " & OrderField & AscDesc &",id desc"
			Set RS = Server.CreateObject("AdoDb.RecordSet")
			RS.Open SQLStr, conn, 1, 1
			If Not RS.Eof Then
			 If Page >1 and (Page - 1) * MaxPerPage < totalPut Then
					RS.Move (Page - 1) * MaxPerPage
			 Else
					Page = 1
			 End If
			 Set IXML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
			 RS.Close : Set RS=Nothing
			 Call showContent()
            End If
		Else
			If DataBaseType=1 Then
					Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
					Set Cmd.ActiveConnection=conn
					Cmd.CommandText="KS_GetPageRecords"
					Cmd.CommandType=4	
					CMD.Prepared = true 
					Cmd.Parameters.Append cmd.CreateParameter("@tblName",202,1,200)
					Cmd.Parameters.Append cmd.CreateParameter("@fldName",202,1,200)
					Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
					Cmd.Parameters.Append cmd.CreateParameter("@pageindex",3)
					Cmd.Parameters.Append cmd.CreateParameter("@ordertype",3)
					Cmd.Parameters.Append cmd.CreateParameter("@strWhere",202,1,1000)
					Cmd.Parameters.Append cmd.CreateParameter("@fieldIds",202,1,1000)
					Cmd("@tblName")=KS.C_S(ChannelID,2)
					Cmd("@fldName")= OrderField
					Cmd("@pagesize")=MaxPerPage
					Cmd("@pageindex")=page
					Cmd("@ordertype")=OrderType
					Cmd("@strWhere")=Param
					Cmd("@fieldIds")=FieldStr
					Set Rs=Cmd.Execute
		   Else
			SQLStr=KS.GetPageSQL(KS.C_S(ChannelID,2),OrderField,MaxPerPage,Page,OrderType,Param,FieldStr)
			Set RS = Server.CreateObject("AdoDb.RecordSet")
			RS.Open SQLStr, conn, 1, 1
		   End If
		   If Not RS.EOF Then
				Set IXML=KS.RSToxml(RS,"row","")
				RS.Close :Set RS=Nothing
				Call showContent()
		   Else
			  RS.Close :Set RS=Nothing
			 .echo "<tr><td colspan=18 align='center' height='35' class='splittd'><font color=red>�Բ���û���ҵ��κ�" &KS.C_S(ChannelID,3) & "!</font></td></tr>"
		   End If
		End If
			  .echo ("<table width='100%' border='0' cellspacing='0' cellpadding='0' align='center'>")
			  .echo ("<tr><td nowrap><b>ѡ��</b><a href='javascript:void(0)' onclick='javascript:Select(0)'>ȫѡ</a>  <a href='javascript:void(0)' onclick='javascript:Select(1)'>��ѡ</a>  <a href='javascript:void(0)' onclick='javascript:Select(2)'>��ѡ</a>")
			  .echo ("</td>")
			  .echo ("<td><td align='right' nowrap>")
			  
		If ComeFrom="RecycleBin" Then
			  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
			  .echo ("<tr><td style='padding-left:20px'>")
			  .echo ("<input type=""button"" value=""������ԭ"" onclick=""BackRecely()"" class=""button"">")
			  .echo (" <input type=""button"" value=""����ɾ��"" onclick=""Delete()"" class=""button"">")
			  .echo (" <input type=""button"" value=""һ�����"" onclick=""DelAll()"" class=""button"">")
			  .echo ("</td></tr>")
			  .echo ("</table>")
		Else
			  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
			  .echo ("<tr><td width='49%' align='center' nowrap>")
			  .echo ("<fieldset align=center><legend>�趨</legend>")
			  .echo ("<input type=""button"" value=""�Ƽ�"" onclick=""set(this,1)"" class=""button"">")
			  .echo (" <input type=""button"" value=""�õ�"" onclick=""set(this,2)"" class=""button"">")
			  .echo (" <input type=""button"" value=""����"" onclick=""set(this,3)"" class=""button"">")
			  .echo (" <input type=""button"" value=""����"" onclick=""set(this,4)"" class=""button"">")
			  .echo (" <input type=""button"" value=""ͷ��"" onclick=""set(this,5)"" class=""button"">")
			  .echo (" <input type=""button"" value=""�̶�"" onclick=""set(this,6)"" class=""button"">")
			  .echo (" <input type=""button"" value=""����"" onclick=""set(this,7)"" class=""button"">")
			  
			  .echo ("</fieldset>")
			  .echo ("</td><td width='1%'>&nbsp;</td><td width='49%' align='center' nowrap>")
			  .echo ("<fieldset align=center ><legend>ȡ��</legend>")
			  .echo ("<input type=""button"" value=""�Ƽ�"" onclick=""set(this,8)"" class=""button"">")
			  .echo (" <input type=""button"" value=""�õ�"" onclick=""set(this,9)"" class=""button"">")
			  .echo (" <input type=""button"" value=""����"" onclick=""set(this,10)"" class=""button"">")
			  .echo (" <input type=""button"" value=""����"" onclick=""set(this,11)"" class=""button"">")
			  .echo (" <input type=""button"" value=""ͷ��"" onclick=""set(this,12)"" class=""button"">")
			  .echo (" <input type=""button"" value=""�̶�"" onclick=""set(this,13)"" class=""button"">")
			  .echo (" <input type=""button"" value=""����"" onclick=""set(this,14)"" class=""button"">")
			  .echo ("</fieldset>")
			  .echo ("</td></tr>")
			  .echo ("</table>")
		  End If
			  
			  .echo ("</td></tr></form></table>")
			  .echo "</div>"
			  .echo ("<table border='0' width='100%'><tr>")
			  If KS.C_S(ChannelID,7)<>0 Then
			  .echo ("<td align='center' width='100'><input class='button' onclick='CreateHtml()' type='button' value='����'> <input class='button' onclick='MoveClass()' type='button' value='�ƶ�'></td>")
			  End If
			  .echo ("<td>")
			  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			  .echo ("</td></tr></table>")

			  
			  .echo ("<table style='clear:both;background:url(images/ico/bottombg.gif);margin-top:5px;' height='43' border='0' width='100%' align='center'>")
			  .echo ("<form action='KS.ItemInfo.asp' name='searchform' method='get'>")
			  .echo ("<tr><td><img src='images/ico/search.gif' align='absmiddle'>������")
			  .echo ("&nbsp;�������� <select name='searchtype'>")
			  If ChannelID=5 Then 
			  If SearchType="4" Then .echo ("<option value=4 selected>��Ʒ���</option>") Else .echo ("<option value=4>��Ʒ���</option>")
			  If SearchType="0" Then .echo ("<option value=0 selected>��Ʒ����</option>") Else .echo ("<option value=0>��Ʒ����</option>")
			  If SearchType="5" Then .echo ("<option value=5 selected>��Ʒ����Ʒ��</option>") Else .echo ("<option value=5>��Ʒ����Ʒ��</option>")
			  Else
			  If SearchType="0" Then .echo ("<option value=0 selected>�ĵ�����</option>") Else .echo ("<option value=0>�ĵ�����</option>")
			  End If
			  If SearchType="1" Then .echo ("<option value=1 selected>�ĵ�¼��Ա</option>") Else .echo("<option value=1>�ĵ�¼��Ա</option>")
			  If SearchType="2" Then .echo ("<option value=2 selected>�ĵ��ؼ���</option>") Else .echo ("<option value=2>�ĵ��ؼ���</option>")
			  If SearchType="3" Then .echo ("<option value=3 selected>�ĵ�����</option>") Else .echo ("<option value=3>�ĵ�����</option>")
			  .echo ("</select> <input type='text' title='�ؼ��ֿ�����' class='textbox' value='" & KeyWord &"' size='12' name='keyword'>&nbsp;<span style='cursor:pointer' onclick='setstatus(this)'><b>�޸����ڣ�</b></span>")
			  If StartDate <> "" And EndDate <> "" Then
			  .echo ("<span id='SearchDate'>��ʼ����<input onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#StartDate')[0]);$('#StartDate').focus();"" type='text' size='12' readonly  name='StartDate' value='" & StartDate & "' style='cursor:pointer'  id='StartDate'>&nbsp;��������<input type='text' readonly size=12 value='" & EndDate & "' name='EndDate' id='EndDate' style='cursor:pointer'  onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#EndDate')[0]);$('#EndDate').focus();""></span>")
			  Else
			  .echo ("<span style='display:none' id='SearchDate'>��ʼ����<input onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#StartDate')[0]);$('#StartDate').focus();"" type='text' size='12' readonly  name='StartDate' style='cursor:pointer'  id='StartDate'>&nbsp;��������<input type='text' readonly size=12 name='EndDate' id='EndDate' style='cursor:pointer'  onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#EndDate')[0]);$('#EndDate').focus();""></span>")
			  End If
			  .echo ("&nbsp;<input type='submit' class='button' value='��ʼ����'><input type='hidden' value='" & ChannelID & "' name='channelid'><input type='hidden' value='" & ComeFrom & "' name='ComeFrom'></td>")
			  .echo ("</tr>")
			  .echo ("</form>")
			  .echo ("</table>")
		  .echo ("</div>")
		  .echo ("</body>")
		  .echo ("</html>")
		  End With
		End Sub

Sub showContent()
    If Not IsObject(IXml) Then Exit Sub
		    Dim ItemIcon,ItemId
			With KS
			For Each INode In IXml.DocumentElement.SelectNodes("row")
			        ItemId=INode.SelectSingleNode("@id").text
					If Not KS.IsNul(INode.SelectSingleNode("@photourl").text) Then
						 ItemIcon="Images/ico/doc1.gif"
					Else
						 ItemIcon="Images/ico/doc0.gif"
					End If
					    AttributeStr = ""
						If KS.ChkClng(INode.SelectSingleNode("@recommend").text) = 1 Or KS.ChkClng(INode.SelectSingleNode("@popular").text) = 1 Or KS.ChkClng(INode.SelectSingleNode("@strip").text) = 1 Or KS.ChkClng(INode.SelectSingleNode("@rolls").text) = 1 Or KS.ChkClng(INode.SelectSingleNode("@slide").text) = 1 Or KS.ChkClng(INode.SelectSingleNode("@istop").text) = 1 Then
								  If KS.ChkClng(INode.SelectSingleNode("@recommend").text) = 1 Then AttributeStr = AttributeStr & (" <span title=""�Ƽ�" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""green"">��</font></span>&nbsp;")
								  If KS.ChkClng(INode.SelectSingleNode("@popular").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""����" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""red"">��</font></span>&nbsp;")
								  If KS.ChkClng(INode.SelectSingleNode("@strip").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""����ͷ��"" style=""cursor:default""><font color=""#0000ff"">ͷ</font></span>&nbsp;")
								  If KS.ChkClng(INode.SelectSingleNode("@rolls").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""����" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""#F709F7"">��</font></span>&nbsp;")
								  If KS.ChkClng(INode.SelectSingleNode("@slide").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""�õ�Ƭ" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""black"">��</font></span>")
								  IF KS.ChkClng(INode.SelectSingleNode("@istop").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""�̶�" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""brown"">��</font></span>")
								  If KS.C_S(Channelid,6)=1 Then
								   IF KS.ChkClng(INode.SelectSingleNode("@isvideo").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""��Ƶ" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""#ff6600"">Ƶ</font></span>")
								  End If
								  If AttributeStr="" Then AttributeStr="---"
					   Else
								AttributeStr = "---"
					   End If
					   
					If KS.ChkClng(KS.G("Status"))=10 Then
					   Dim RSS,HasSignUser,XML,Node,MustSignUserArr,SignUser,NoSignUser,S,AttrStr
					   Set RSS=Conn.Execute("Select top 500 username From KS_ItemSign Where ChannelID=" & ChannelID & " and infoid=" & itemId)
					   If Not RSS.EOf Then
						   SET xml=KS.RsToXml(RSS,"row","")
						   for each node in xml.documentelement.selectnodes("row")
							 if HasSignUser="" then 
							   HasSignUser=node.selectSingleNode("@username").text
							 else
							   HasSignUser=HasSignUser& "," & node.selectSingleNode("@username").text
							 end if
						   next
					   End If
					   RSS.Close
					   
					   SignUser=INode.SelectSingleNode("@signuser").text :  NoSignUser="" : MustSignUserArr=Split(SignUser,",")
					   If IsArray(MustSignUserArr) Then

					   For S=0 To Ubound(MustSignUserArr)
						  If KS.FoundInArr(HasSignUser,MustSignUserArr(S),",")=false Then
							if NoSignUser="" then
							  NoSignUser=MustSignUserArr(S)
							else
							  NoSignUser=NoSignUser & "," & MustSignUserArr(S)
							end if
						  End If
					   Next
					   End If
					   If NoSignUser="" Then AttrStr="<font color=blue>ǩ�����</font>" Else AttrStr="<font color=red>ǩ����...</font>"
					   TitleStr =" title='��ǩ���û�:" & HasSignUser & "&#13;&#10;δǩ���û�:"& NoSignUser &"'"
					Else
                     TitleStr = " TITLE='�� ��:" & INode.SelectSingleNode("@title").text & "&#13;&#10;�� ��:" & INode.SelectSingleNode("@adddate").text & "&#13;&#10;¼ ��:" & INode.SelectSingleNode("@inputer").text & "'"
					End If
						.echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & ItemId & "' onclick=""chk_iddiv('" & ItemId & "')"">")
						.echo ("<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & ItemId & "')"" type='checkbox' id='c"& ItemId & "' value='" &ItemId & "'></td>")
						 If ChannelID=8 Then
							 .echo ("<td align=""center"" class='splittd'>" & KS.GetGQTypeName(INode.SelectSingleNode("@typeid").text) & "</td>")
						 End If
							 .echo ("<td" & TitleStr & " class='splittd' nowrap><span onDblClick=""ViewArticle(" & INode.SelectSingleNode("@id").text &")"">")
							 .echo ("<a href='javascript:ViewArticle(" & ItemId & ");'><img src=" & ItemIcon & " border=0 align=absmiddle title='Ԥ��'></a>")
							 .echo ("<span style=""cursor:default""><a href='?ID=" & INode.SelectSingleNode("@tid").text &"&channelid=" & ChannelID&"'>[" & KS.C_C(INode.SelectSingleNode("@tid").text,1) &"]</a> "& KS.Gottopic(INode.SelectSingleNode("@title").text,27)) & AttrStr
							 If ChannelID=5 Then
							  if INode.SelectSingleNode("@producttype").text="2" Then .echo "<span title='�Ǽ۳���' style='color:green'> ��</span>" Else If INode.SelectSingleNode("@producttype").text="3" Then .echo "<span style='color:red' title='���۳���'> ��</span>"
							 End If
							 .echo ( "</span></span></td>")						
						
							 
						If IsArray(XmlFieldArr) Then
							 For Fi=0 To Ubound(XmlFieldArr)
							   .echo ("<td class='splittd' nowrap align='center'>&nbsp;")
							   select case lcase(Split(XmlFieldArr(fi),"|")(1))
							    case "modeltype" .echo KS.C_S(ChannelID,3)
								case "attribute" .echo AttributeStr
								case "adddate" .echo FormatDateTime(INode.SelectSingleNode("@adddate").text,2) 
								case "refreshtf" 
								  If KS.C_S(ChannelId,7)="0" then
								     .echo "<span style='color:blue;cursor:default' title='��ģ��û���������ɾ�̬HTML,��������'>��������</span>"
								  Else
								   if INode.SelectSingleNode("@refreshtf").text="1" then
								     .echo "<font color=green>������</font>"
								   else 
								     .echo "<font color='#ff3300'>δ����</font>"
								   end if
								  End If
								case else
							   .echo INode.SelectSingleNode("@" &lcase(Split(XmlFieldArr(fi),"|")(1))).text
							  end  select
							  .echo ("&nbsp;</td>")
							 Next
							 if ComeFrom<>"" Then
							  .Echo ("<td width='60'>" & GetStatus(INode.SelectSingleNode("@verific").text) & "</td>")
							 End If
						Else
							 .echo ("<td align=""center"" class='splittd'>" & INode.SelectSingleNode("@inputer").text & "</td>")
							 .echo ("<td align=""center"" class='splittd'>" & FormatDateTime(INode.SelectSingleNode("@adddate").text,2) & "</td>")
							 .echo ("<td align=""center"" class='splittd'>" & KS.C_S(ChannelID,3) & "</td>")
							 .echo ("<td align=""center"" class='splittd'>" & AttributeStr & "</td>")
							 .echo ("<td align=""center"" class='splittd'>")
							 
							  If ComeFrom="" Then
							    .echo INode.SelectSingleNode("@hits").text
							  Else
							    .echo GetStatus(INode.SelectSingleNode("@verific").text)
							  End If
							 .echo ("</td>")
					End If	 
							 .echo ("<td align=""center"" nowrap class='splittd'>")
							 If ComeFrom="RecycleBin" Then
							 .echo("<a href='?Page=" & Page & "&Action=RecelyBack&" &SearchParam&"&ID=" & ItemId & "'>��ԭ</a> | <a href=""?Action=Delete&Page=" & Page & "&" & SearchParam & "&ID=" & ItemId & """ onclick=""return (confirm('�˲��������棬ȷ������" & KS.C_S(ChannelID,3) & "����ɾ����?'))"">��ɾ</a>")
							 ElseIf ComeFrom="Verify" Then
							  If Cint(INode.SelectSingleNode("@verific").text) =2  Then
							  .echo "<font color=#cccccc>���������</font>"	  
							  Else
								 If Cint(INode.SelectSingleNode("@verific").text) <>3  Then   '����˻�ݸ����²��������
								   .echo "  <a href=""#""  onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=" & ChannelID & "&ComeFrom=Verify&ButtonSymbol=AddInfo&OpStr=" & server.URLEncode(KS.C_S(ChannelID,3) & "���� >> <font color=red>ǩ�ջ�Ա" & KS.C_S(ChannelID,3)) & "</font>';location.href='" & ItemManageUrl & "?ChannelID=" & ChannelID & "&Page=" & Page & "&Action=Verify&ID="&ItemId&"';"">���</a>"
								  If Cint(INode.SelectSingleNode("@verific").text)<>2 Then
								 .echo "&nbsp;<a onclick=""ProcessTuigao(event," & ItemId & ")"" href='#'>�˸�</a>"
								  End IF
								 End If
								 .echo (" <a href=""?Action=Recely&Page=" & Page & "&" & SearchParam & "&ID=" & ItemId & """ onclick=""return (confirm('ȷ������" & KS.C_S(ChannelID,3) & "�������վ��?'))"">����վ</a>")
								End If
							 Else
							 .echo (" <a href='" & ItemManageUrl & "?Page=" & Page & "&Action=Edit&" &SearchParam&"&ID=" & ItemId & "' onclick='parent.frames[""BottomFrame""].location.href=""KS.Split.asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("�༭" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & ID & """;'>�޸�</a> | <a href=""?Action=Recely&Page=" & Page & "&" & SearchParam & "&ID=" & ItemId & """ onclick=""return (confirm('ȷ������" & KS.C_S(ChannelID,3) & "�������վ��?'))"">����վ</a>")
							 If ChannelID=5 Then
							 .echo (" | <a href='KS.ProImages.asp?ProID=" &ItemId &"&ChannelID=5'>ͼƬ</a>")
							 End If
							 End If
							 .echo ("</td>")
							 .echo ("</tr>")
			  Next

			  .echo ("</table>")
			End With
		End Sub
      Function GetStatus(verific)
	     Select Case Cint(verific)
			Case 0: GetStatus = "<span style='color:red'>����</span>"
			Case 1: GetStatus = "<span style='color:blue'>����</span>"
            Case 2: GetStatus = "<span style='color:#999999'>�ݸ�</span>"
            Case 3: GetStatus = "<span style='color:green'>�˸�</span>"
         End Select
	  End Function

      Sub ShowClassList(ChannelID,ID)
		 If KS.S("ComeFrom")<>"" Then Exit Sub
		 
		 With KS
		 '============���Ӽ��书��=======================================
		 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
		 classExtStatus=request.cookies("classExtStatus")
		 if classExtStatus="" Then classExtStatus=1
		 If classExtStatus=1 Then 
		  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
		 Else 
		  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
		 End If
		 '=========================================================----
		 .echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 2px;' ><img src='images/kszk.gif' align='absmiddle'></div>"
		 .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;line-height:30px;margin:5px 1px;border:1px solid #DEEFFA;background-color:#F7FBFE'>"
		 .echo "<div style='padding-top:2px;cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px; top: 2px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='images/close.gif' align='absmiddle'></div>"
		
		Dim P,RSC,Img,j,N,I,XML,Node
		P=" where ClassType=1 and ChannelID=" & ChannelID
		If ID=0 Then
		  P=P & " And tj=1"
		 Img="domain.gif"
		Else
		 P=P & " And TN='" & ID & "'"
		 Img="Smallfolder.gif"
		End If

		 on error resume next
		 Dim ParentID:ParentID = conn.Execute("Select TN From KS_Class  Where ID='" & ID & "'")(0)

		Set RSC=Conn.Execute("select id,foldername,adminpurview from ks_class " & P& " order by root,folderorder")
		If Not RSC.Eof Then 
		 Set XML=.RsToXml(RSC,"row","xmlroot")
		 RSC.Close:Set RSC=Nothing
		 If IsObject(XML) Then
		   If ID<>"0" Then
		    .echo "<a href='?ChannelID=" & ChannelID & "&ID=" & ParentID & "'><img src='images/folder/Back.gif' border=0 align='absmiddle'></a> "
		   End if
		   For Each Node In XML.DocumentElement.SelectNodes("row")
		    If KS.C("SuperTF")=1 or KS.FoundInArr(Node.SelectSingleNode("@adminpurview").text,KS.C("AdminName"),",") or Instr(KS.C("ModelPower"),KS.C_S(ChannelID,10)&"1")>0 Then 
		    .echo "<li style='margin:5px;float:left;width:100px'><img src='images/folder/" & Img & "' align='absmiddle'><a href='?ChannelID=" & ChannelID & "&ID=" & Node.SelectSingleNode("@id").text & "' title='" & Node.SelectSingleNode("@foldername").text & "'>" & .Gottopic(Node.SelectSingleNode("@foldername").text,8) & "</a></li>"
		    End If
		   Next
		 End If
		Else
		  If err Then
		   .echo "<img src='images/folder/AddFolder.gif' align='absmiddle'>����<a href='#' onclick=""location.href='KS.Class.asp?Action=Add&ChannelID=" & ChannelID & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=" & Server.URLEncode("��Ŀ���� >> <font color=red>�����Ŀ</font>") & "';"">�����Ŀ</a>"
		  Else
		   .echo "<a href='?ChannelID=" & ChannelID & "&ID=" & ParentID & "'><img src='images/folder/Back.gif' border=0 align='absmiddle'></a> <a href='#' onclick='CreateNews()'><font color=#4169E1><strong>���" & KS.C_S(Channelid,3) & "</strong></font></a>"
		   End If
		End If
		 .echo "</div>"
		 .echo "<div style=""clear:both""></div>"
		 End With
		End Sub
		
		Sub ShowChannelList()
		  With KS
			 '============�����书��=======================================
			 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
			 classExtStatus=request.cookies("classExtStatus")
			 if classExtStatus="" Then classExtStatus=1
			 If classExtStatus=1 Then 
			  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
			 Else 
			  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
			 End If
			 '=========================================================----
			 .echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 2px;' ><img src='images/kszk.gif' align='absmiddle'></div>"
			 .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;line-height:30px;margin:5px 1px;border:1px solid #DEEFFA;background:#F7FBFE'>"
			 .echo "<div style='padding-top:2px;cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px; top: 2px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='images/close.gif' align='absmiddle'></div>"
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				   .echo "<li style='margin:5px;float:left;width:100px'><img src='images/folder/domain.gif' align='absmiddle'><a href='?ChannelID=" & Node.SelectSingleNode("@ks0").text & "&ComeFrom=RecycleBin' title='" & Node.SelectSingleNode("@ks1").text & "'>" & .Gottopic(Node.SelectSingleNode("@ks1").text,8) & "(<span style='color:red'>" & Conn.Execute("Select Count(ID) From " & Node.SelectSingleNode("@ks2").text & " Where Deltf=1")(0) & "</span>)</a></li>"
			    End If
			next
			.echo "</div>"
			.echo "<div style=""clear:both""></div>"
         End With
		End Sub
		
		'����JS
		Sub AddToJS()
		    DIM JSNameList,JSObj,NewsID
			NewsID=Trim(Request("NewsID"))
			 Set JSObj=Server.CreateObject("Adodb.Recordset")
			 JSObj.Open "Select JSName,JSID From KS_JSFile Where JSType=1 And JSConfig NOT LIKE 'GetExtJS%'",Conn,1,1
			 IF NOT JSObj.EOF THEN
				 JSNameList="<Option Value='0'></Option>"
			  DO While NOT JSObj.EOF 
				 JSNameList=JSNameList & "<Option value=" & JSObj("JSID") &">" & Trim(JSObj("JSName")) & "</Option>"
				 JSObj.MoveNext
			  LOOP
			 Else
				 JSNameList=JSNameList & "<Option value=0>---��û�н�����JS---</Option>"
			 END IF
			JSObj.Close:Set JSObj=Nothing
			%>  
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; chaRSet=gb2312">
			<title>��������JS</title>
			<link href="Include/Admin_Style.css" rel="stylesheet">
			<link href="Include/ModeWindow.css" rel="stylesheet">
			<script language="JavaScript" src="../KS_Inc/common.js"></script>
			</head>
			<body topmargin="0" leftmargin="0" scroll=no>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <form name="myform" action="?ChannelID=<%=ChannelID%>&Action=JS" method="post">
			  <input type="hidden" value="Add" Name="Flag">
			  <input type="hidden" name="JSName">
			  <input type="hidden" value="<%=NewsID%>" Name="NewsID"> 
			  <tr> 
				<td height="18">&nbsp;</td>
			  </tr>
			  <tr> 
				<td height="30" align="center"> <strong>��ѡ��JS����</strong> 
				  <select name="JSID">
					  <%=JSNameList%>
				  </select>
				</td>
			  </tr>
			  <tr align="center"> 
				<td height="30"> <input type="button" class="button" name="button1" value="����JS" onClick="CheckForm()"> 
				  &nbsp; <input type="button" class="button" onClick="window.close();" name="button2" value=" ȡ�� "> 
				</td>
			  </tr>
			  </form>
			</table>
			</body>
			</html>
			<Script>
			function CheckForm()
			{
			 if (document.myform.JSID.value=='0')
			  { alert('�Բ���,��û��ѡ��JS����!');
				 document.myform.JSID.focus();
				 return false;
			  }
			  document.myform.JSName.value=document.myform.JSID.options[document.myform.JSID.selectedIndex].text
			  document.myform.submit();
			  return true
			}
			</Script> 
			<%IF Request.Form("Flag")="Add" Then
			   Dim RS,OldJSID,JSID,NewsIDArr,K
			   JSID=Trim(Request.Form("JSID"))
			   NewsIDArr=Split(NewsID,",")
			   Set RS=Server.CreateObject("Adodb.RecordSET")
			   For K=Lbound(NewsIDArr) To Ubound(NewsIDArr)
				  RS.Open "Select JSID From " & KS.C_S(ChannelID,2) &" Where ID=" & NewsIDArr(K),Conn,1,3
				  IF  Not RS.Eof THEN
						 OldJSID=Trim(RS("JSID"))
					   IF Trim(RS(0))="0" or Trim(RS(0))="" or isnull(RS(0)) Then
						  RS(0)=JSID & ","
					   Elseif InStr(OldJSID,JSID)=0 then
						  RS(0)=RS(0) & JSID & ","
					   End if
					   RS.UPDate
					   
					 End IF
                  RS.Close
			   Next
			            'ˢ��JS
					   Dim KSRObj,JSName
					   JSName=Trim(Request.Form("JSName"))
					   Set KSRObj=New Refresh
					   KSRObj.RefreshJS(JSName)
					   Set KSRObj=Nothing
			   Set RS=Nothing
			   KS.Echo "<script>alert('�����ɹ�!');window.close();</script>"
			End IF
		End Sub
		
		'�����ƶ�
		Sub MoveClass()
		Dim RS, IDArr, K
		 Dim ID:ID=KS.FilterIDs(Request("ID"))
		 If id="" Then KS.Die "<script>alert('������!');parent.location.reload();</script>" 
		 Dim ChannelID:ChannelID=KS.ChkClng(Request("ChannelID"))
		 If ChannelID=0 Then ChannelID=1
		 If KS.G("Flag")="save" Then
		   Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set Tid='" & KS.G("Tid") & "' Where ID in(" & id &")")
		   Conn.Execute("Update KS_ItemInfo Set Tid='" & KS.G("Tid") & "' Where ChannelID=" & ChannelID & " and InfoID in(" & ID & ")")
		   KS.Die "<script>alert('��ϲ�������ƶ���Ŀ����Ŀ!');parent.location.reload();</script>" 
		 End If
		 %>
		 	<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; chaRSet=gb2312">
			<title>��������</title>
			<link href="Include/Admin_Style.css" rel="stylesheet">
			</head>
			<body>
			 <br/>
			 <form name="myform" method="post" action="KS.ItemInfo.asp">
			 <input type="hidden" name="channelid" value="<%=channelid%>"/>
			 <input type="hidden" name="id" value="<%=id%>"/>
			 <input type="hidden" name="action" value="MoveClass"/>
			 <input type="hidden" value="save" name="flag"/>
			 <div style="text-align:center">
			 <strong>��ѡ�е������ƶ�����Ŀ</strong> <select name="tid">
			 <%=KS.LoadClassOption(ChannelID,true)%>
			 </select>
			 <input type="submit" value="ȷ���ƶ�" class="button">
			 </div>
			 </form>
			 <br/>
			 <span style='color:blue;padding-left:10px;'>
			 Tips:���������վ���������ɾ�ҳHTML���ܣ������ƶ���������������Ӧ����Ŀ��
			 </span>
			</body>
			</html>
		<%
		End Sub
		
		'��������
		Sub SetAttribute()
		 Dim RS, IDArr, K
		 Dim ID:ID=Trim(Request("ID"))
		 Dim ChannelID:ChannelID=KS.ChkClng(Request("ChannelID"))
		 If ChannelID=0 Then ChannelID=1
		 %>
		 	<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; chaRSet=gb2312">
			<title>��������</title>
			<link href="Include/Admin_Style.css" rel="stylesheet">
			<script language="JavaScript" src="../KS_Inc/common.js"></script>
			<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
	        <script src="images/pannel/tabpane.js" language="JavaScript"></script>
	        <link href="images/pannel/tabpane.CSS" rel="stylesheet" type="text/css">
				 <script language="javascript">
				  $(document).ready(function(){
				   $("#channelids").change(function(){
					 if ($(this).val()!=0){
					  $(parent.frames["FrameTop"].document).find("#ajaxmsg").toggle();
					  $.get("../plus/ajaxs.asp",{action:"GetClassOption",channelid:$(this).val()},function(data){
						 $(parent.frames["FrameTop"].document).find("#ajaxmsg").toggle();
						 $("select[name=ClassID]").empty();
						 $("select[name=ClassID]").append(unescape(data));
						 $("input[name=ChannelID]").val($("#channelids").val());
						 if ($("input[name=ChannelID]").val()==5 || $("input[name=ChannelID]").val()==7 || $("input[name=ChannelID]").val()==8){
						  $("#showauthor").hide();
						  $("#showorigin").hide();
						  }else{
						  $("#showauthor").show();
						  $("#showorigin").show();
						  }
					   });
					 }
				   });
				  })

				function SelectAll(){
				  $("select[name=ClassID]>option").each(function(){
				   $(this).attr("selected",true);
				  })
				}
				function UnSelectAll(){
				  $("select[name=ClassID]>option").each(function(){
				   $(this).attr("selected",false);
				  })
				}
				</SCRIPT>			
           </head>
			<body topmargin="0" leftmargin="0" scroll="no">
			<div class="topdashed sort">���������ĵ�����</div>
			<div style="height:94%; overflow: auto; width:100%">
			<iframe src="about:blank" width="0" height="0" name="_hiddenframe" id="_hiddenframe" style="display:none"></iframe>
			<table width="100%" style="margin-top:10px" border="0" align="center"  cellspacing="1" class='ctable'>
			<form name="myform" action="?Action=SetAttribute" method="post" target="_hiddenframe">
			  <input type='hidden' name='ChannelID' id='ChannelID' value='<%=ChannelID%>'>
			  <input type="hidden" value="Add" Name="Flag">
			  <tr class='tdbg' id='choose2'<%if ID<>"" then response.write " style='display:none'"%>>
				<td valign='top' rowspan='100' width='200'>
				<font color=red>��ʾ��</font>���԰�ס��Shift��<br />��Ctrl�������ж����Ŀ��ѡ��<br />
				<%if ChannelID<>5 then%>
				<select id='channelids' name='channelids' style='width:200px'>
				 <option value='0'>---��ѡ��ģ��---</option>
				 <%
				If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				  Response.write "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
				 End If
				next
				%>
				</select>
				<%end if%>
				
			<Select style="WIDTH: 200px; HEIGHT: 380px" multiple size=2 name="ClassID">
			 <%=KS.LoadClassOption(ChannelID,false)%>
			</Select>
			<div align=center>
			   <Input onclick=SelectAll() type=button class="button" value="ѡ��������Ŀ" name=Submit><br />
			   <Input onclick=UnSelectAll() type=button value="ȡ��ѡ����Ŀ" class="button" name=Submit></div>
                </td>
			  </tr>
			  <tr class='tdbg'>
			     <TD valign="top">
				 
				   
				        <table border="0" width="100%" cellpadding="0" cellspacing="1" class="ctable">
				            <tr>
							 <td class='clefttitle' align='right'><strong>����ѡ��:</strong></td>
							 <td><input type=radio name=choose value='0'<%if ID<>"" then response.write" checked"%> onClick="choose1.style.display='';choose2.style.display='none';"> ���ĵ�ID&nbsp;&nbsp;		<input type=radio name=choose value='1' onClick="choose2.style.display='';choose1.style.display='none';"<%if ID="" then response.write " checked" else response.write "disabled"%>> ���ĵ�����</td>
						  </tr>
						  <tr class='tdbg' id='choose1'<%if ID="" then response.write " style='display:none'"%>>
							 <td class='clefttitle' align='right'><strong>�ĵ�ID��</strong>���ID���á�,���ֿ�</td>
							 <td><input type='text' size='50' value='<%=ID%>' name='ID'></td>
						  </tr>
						</table>
						
						
				  <%if ChannelID=5 then%>
				    <script type="text/javascript">
					 function setPrice(p)
					 {
					   $("#groupprice").find("input").each(function(){
					     $(this).val(p);
					   });
					 
					   $("input[name='DiscountPriceMarket']").val(p);
					   $("input[name='DiscountPrice']").val(p);
					   $("input[name='DiscountPriceMember']").val(p);
					   $("input[name='DiscountScore']").val(p);
					 }
					function regInput(obj, reg, inputStr)
					{
						var docSel = document.selection.createRange()
						if (docSel.parentElement().tagName != "INPUT")    return false
						oSel = docSel.duplicate()
						oSel.text = ""
						var srcRange = obj.createTextRange()
						oSel.setEndPoint("StartToStart", srcRange)
						var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
						return reg.test(str)
					}
					</script>		
					<div class="tab-page" id="SetAttrPanel">
						<SCRIPT type=text/javascript>
							   var tabPane1 = new WebFXTabPane( document.getElementById( "SetAttrPanel" ), 1 )
						</SCRIPT>
								 
					<div class=tab-page id=price-page1>
					<H2 class=tab>��������</H2>
					<SCRIPT type=text/javascript>
						 tabPane1.addTabPage( document.getElementById( "price-page1" ) );
					</SCRIPT>
								
					 <table border="0" width="100%" cellpadding="0" cellspacing="1">
						  <tr class='tdbg'> 
						    
                            <td class='clefttitle' align='right' width="80"><strong>�ۿ۱�־:</strong></td>
							<td class='clefttitle' height='25' style="text-align:left">
							<label onClick="$('#zkl').hide()"><input name='ProductType' type='radio' value='0' checked>��������</label>
							<br/>
							<label onClick="$('#zkl').show();setPrice(10)"><input name='ProductType' type='radio' value='1'>ȫ����λ��ԭʼ���ۼ�</label>
							<br/>
							<label style="color:green" onClick="$('#zkl').show();setPrice(11)"><input name='ProductType' type='radio' value='2'>��Ϊ�Ǽ�����</label>
							<br/>
							<label style="color:red" onClick="$('#zkl').show();setPrice(9.8)"><input name='ProductType' type='radio' value='3'>��Ϊ��������</label>
							<div id='zkl' style="display:None">
							
							  <table border="0" width="100%" cellpadding="0" cellspacing="1">
							   <tr>
							    <td class='clefttitle' height='25' align='left' nowrap>
								<label><input type='checkbox' name='ePriceMarket' value='1'><strong>�� �� ��:</strong></label>
								</td>
								<td>
							 <div>��<font color="#FF0000">��ԭʼ���ۼۡ�</font>Ϊ��׼,��<font color="blue">���г��ۡ�</font>��<input size="4" style="text-align:center" name='DiscountPriceMarket' type='text' value='9' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">������ѡ�е�������Ʒ</div>
							     <font color="#999999">Tips:�г��۽����ڸ��ͻ��ο��ļ۸񣬲����ڽ��ס�</font>
							    </td>
							   </tr>
							   <tr><td colspan=2><hr color="green" size=1></td></tr>
							   <tr>
							    <td class='clefttitle' height='25' align='left' nowrap>
								<label><input type='checkbox' name='ePrice' value='1'><strong>�ǻ�Ա��:</strong></label>
								</td>
								<td>
							 <div>��<font color="#FF0000">��ԭʼ���ۼۡ�</font>Ϊ��׼,��<font color="blue">����ǰ���ۼۡ�</font>��<input size="4" style="text-align:center" name='DiscountPrice' type='text' value='9' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">������ѡ�е�������Ʒ</div>
							     <font color="#999999">Tips:��ǰ���ۼ�ָ���Ƿ�ע���Ա�����ļ۸�</font>
							    </td>
							   </tr>
							   <tr><td colspan=2><hr color="green" size=1></td></tr>
							   <tr>
							    <td class='clefttitle' height='25' align='left'>
								 <label><input type='checkbox' name='ePriceMember' value='1'><strong>�� Ա ��:</strong></label>
								 </td>
								<td>
							 <div>��<font color="#FF0000">��ԭʼ���ۼۡ�</font>Ϊ��׼
							 <br/>
							 1�������Ʒ��<font color="#FF0000">����Աͳһָ���ۡ�</font>,��<font color="blue">����Ա�ۡ�</font>��<input size="4" style="text-align:center" name='DiscountPriceMember' type='text' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" value='9'>������ѡ�е�������Ʒ
							 <br/>
							 2�������Ʒ��<font color="#FF0000">����ϸ���û�Ա�ۡ�</font> ,��
							 <br/>
							 <%
							  Response.Write "<table border='0' id='groupprice' width='80%'>"
							  Response.Write "<tr class='clefttitle'><td align='center'><b>��Ա��</b></td><td align='center'><b>�۸�</b></td></tr>"
							  Call KS.LoadUserGroup()
							  For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
							  Response.Write "<tr><td>" &Node.SelectSingleNode("@groupname").text & "</td><td>��<input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' style='text-align:center' size='4' name='price" & Node.SelectSingleNode("@id").text & "'  value='9' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))""> ������ѡ�е�������Ʒ</td></tr>"
							 Next
							  Response.Write "</table>"

							 %>
							 
							 </div>
							    </td>
							   </tr>
							   <tr><td colspan=2><hr color="green" size=1></td></tr>
							   <tr>
							    <td class='clefttitle' height='25' align='left' nowrap>
								<label><input type='checkbox' name='eScore' value='1'><strong>��������:</strong></label>
								</td>
								<td>
							 <div>��<font color="#FF0000">��ԭʼ���ۼۡ�</font>Ϊ��׼,��<font color="blue">���������֡�</font>��<input size="4" style="text-align:center" name='DiscountScore' type='text' value='9' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">������ѡ�е�������Ʒ</div>
							     <font color="#999999">Tips:��������ָ�ͻ��ɹ�������Ʒ�ɵõ������ͻ��֡�</font>
							    </td>
							   </tr>
							   
							  </table>
							
							</div>
							
							
							</td>
						  </tr>
	                 </table>
					   </div>
					   
					   
					   <div class=tab-page id=kbxs-page1>
					<H2 class=tab>��ʱ����</H2>
					<SCRIPT type=text/javascript>
						 tabPane1.addTabPage( document.getElementById( "kbxs-page1" ) );
					</SCRIPT>
								
					 <table border="0" width="100%" cellpadding="0" cellspacing="1">
					     <tr class='tdbg'>
						  <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eLimitBuy' value='1'></td>
						 <%
						 with response
							.Write "  <td class='clefttitle' align='right'><strong><font color=green>�Ƿ���ʱ����:</font></strong></td>"
							.Write "  <td style='padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>"
							.Write "<label onclick=""$('#LimitBuy').hide();""><input name='IsLimitbuy' type='radio'  value='0' checked> ��������</label> &nbsp;&nbsp;<label onclick=""$('#LimitBuy').show();$('#LimitBuyTaskID1').show();$('#LimitBuyTaskID2').hide();""><input name='IsLimitbuy' type='radio'  value='1'> ��ʱ����</label>&nbsp;&nbsp;<label onclick=""$('#LimitBuy').show();$('#LimitBuyTaskID1').hide();$('#LimitBuyTaskID2').show();""><input name='IsLimitbuy' type='radio'  value='2'> ��������</label>"
							.Write "<div id='LimitBuy' style='margin-tio:10px;padding:10px;display:none;border:0px solid #ff6600'>"

							
							.Write "��������:"
							.Write "<select name='LimitBuyTaskID1' id='LimitBuyTaskID1' style='display:none'>"
							.Write "<option value=''>---��ѡ��---</option>"
							
							 Dim RST:Set RST=Conn.Execute("Select ID,taskname from KS_ShopLimitBuy Where TaskType=1 and Status=1 Order by id desc")
							 Do While NOt RST.Eof
								.Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
							 RST.MoveNext
							 Loop
							 RST.CLose 
							 .Write "</select>"
							 .Write "<select name='LimitBuyTaskID2' id='LimitBuyTaskID2' style='display:none'>"
							.Write "<option value=''>---��ѡ��---</option>"
					
							 
							 Set RST=Conn.Execute("Select ID,taskname from KS_ShopLimitBuy Where TaskType=2 and Status=1 Order by id desc")
							 Do While NOt RST.Eof
								.Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
							 RST.MoveNext
							 Loop
							  RST.Close: Set RST=Nothing
							  .Write "</select>"
							 
							.Write " <br/>"
							.Write "�� �� ��:<input type='text' style='text-align:center' name='LimitBuyPrice' value='100' size='6'  value='100' size='4' maxlength='4' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox'>Ԫ<br/>"
							.Write "��������:<input type='text' name='LimitBuyAmount' id='LimitBuyAmount' value='100' size='10'/>��   ������������������Ʒ��<br/>"
							.Write "</div>"
							.Write "</td>"
							.Write "</tr>"
		              End With
						 
						 %>
					 </table>
					</div>
					   
								
					<div class=tab-page id=att-page1>
					<H2 class=tab>��������</H2>
					<SCRIPT type=text/javascript>
						 tabPane1.addTabPage( document.getElementById( "att-page1" ) );
					</SCRIPT>
				<%end if%>

						
						<table border="0" width="100%" cellpadding="0" cellspacing="1">
						  <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eTemplateID' value='1'></td>
							<td class='clefttitle' align='right'><strong>�ĵ�Webģ��:</strong></td>
							<td><input type="text" size='40' name='TemplateID' id='TemplateID' class='textbox'>&nbsp;<%=KSCls.Get_KS_T_C("$('#TemplateID')[0]")%></td>
						  </tr>
						  <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eWapTemplateID' value='1'></td>
							<td class='clefttitle' align='right'><strong>�ĵ�Wapģ��:</strong></td>
							<td><input type="text" size='40' name='WapTemplateID' id='WapTemplateID' class='textbox'>&nbsp;<%=KSCls.Get_KS_T_C("$('#WapTemplateID')[0]")%></td>
						  </tr>
						  <tr class='tdbg'> 
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eKeyWords' value='1'></td>
							<td class='clefttitle' align='right'><strong>�� �� ��:</strong></td>
							<td><input type="text" size='40' name='KeyWords' id='KeyWords' class='textbox'>&nbsp; <select name='SelKeyWords' style='width:100px' onChange='InsertKeyWords($("#KeyWords")[0],this.options[this.selectedIndex].value)'>
					<option value="" selected> </option><option value="Clean" style="color:red">���</option>"
					<%=KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")%>
					</select></td>
						  </tr>
						  <tr class='tdbg' id='showauthor'<%If ChannelID=5 or ChannelID=7 or ChannelID=8 Then KS.Echo " style='display:none'"%>> 
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eAuthor' value='1'></td>
							<td class='clefttitle' align='right'><strong>�ĵ�����:</strong></td>
							<td> <input name='author' type='text' id='author' size=20 class='textbox'><<��<font color='blue'><font color='#993300' onclick='$("#author").val("δ֪")' style='cursor:pointer;'>δ֪</font></font>����<font color='blue'><font color='#993300' onclick="$('#author').val('����')" style='cursor:pointer;'>����</font></font>��
							<select name='SelAuthor' style='width:100px' onChange="$('#author').val(this.options[this.selectedIndex].value)">")
						<option value="" selected> </option><option value="" style="color:red">���</option>
						<%=KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=1 and OriginType=1 Order BY AddDate Desc")%>
						 </select></td>
						  </tr>
						  <tr class='tdbg' id='showorigin'<%If ChannelID=5 or ChannelID=7 or ChannelID=8 Then KS.Echo " style='display:none'"%>>
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eOrigin' value='1'></td>
							<td class='clefttitle' align='right'><strong>�ĵ���Դ:</strong></td>
							<td nowrap><input name='Origin' id='Origin' type='text' size=20 class='textbox'><<��<font color='blue'><font color='#993300' onclick="$('#Origin').val('����')" style='cursor:pointer;'>����</font></font>����<font color='blue'><font color='#993300' onclick="$('#Origin').val('��վԭ��')" style='cursor:pointer;'>��վԭ��</font></font>����<font color='blue'><font color='#993300' onclick="$('#Origin').val('������')" style='cursor:pointer;'>������</font></font>��
						<select name='selOrigin' style='width:100px' onChange="$('#Origin').val(this.options[this.selectedIndex].value)">
						<option value="" selected> </option><option value="" style="color:red">���</option>
						<%=KSCls.Get_O_F_D("KS_Origin","OriginName","OriginType=0 Order BY AddDate Desc")%>
						</select></td>
						</tr>
						 <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='erank' value='1'></td>
							<td class='clefttitle' align='right'><strong>�ĵ��ȼ�:</strong></td>
							<td><select name='rank'>
							 <option>��</option>
							 <option>���</option>
							 <option selected>����</option>
							 <option>�����</option>
							 <option>������</option>
							</select>
						   </td>
						  </tr>		
						 <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='ehits' value='1'></td>
							<td class='clefttitle' align='right'><strong>���������:</strong></td>
							<td><input type='text' value='0' name='hits' size='5'>�� <font color=#777777>˵����ԭ��������ۼ�</font></td>
						  </tr>		
						 <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eAdddate' value='1'></td>
							<td class='clefttitle' align='right'><strong>���ʱ��:</strong></td>
							<td><input type='text' value='<%=now%>' name='AddDate' size='20'> <font color=#777777>��ʽ:2008-12-1 10:10</font></td>
						  </tr>		
						  <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eRecommend' value='1'></td>
							<td class='clefttitle' align='right'><strong>�Ƿ��Ƽ�:</strong></td>
							<td><input name='Recommend' type='radio' id='Recommend' value='1'> ��  <input name='Recommend' type='radio' id='Recommend' value='0' checked> ��</td>
						 </tr>
						 <tr class='tdbg'>
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eIsTop' value='1'></td>
							<td class='clefttitle' align='right'><strong>�Ƿ�̶�:</strong></td>
							<td><input name='IsTop' type='radio' value='1'> ��  <input name='IsTop' type='radio' value='0' checked> ��</td>
						</tr>
						<tr class='tdbg'>
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eRolls' value='1'></td>
							<td class='clefttitle' align='right'><strong>�Ƿ����:</strong></td>
							<td><input name='Rolls' type='radio' value='1'> ��  <input name='Rolls' type='radio' value='0' checked> ��</td>
					   </tr>
					   <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='ePopular' value='1'></td>
							<td class='clefttitle' align='right'><strong>�Ƿ�����:</strong></td>
							<td><input name='Popular' type='radio' value='1'> ��  <input name='Popular' type='radio' value='0' checked> ��</td>
					  </tr>
					  <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eStrip' value='1'></td>
							<td class='clefttitle' align='right'><strong>�Ƿ�ͷ��:</strong></td>
							<td><input name='Strip' type='radio' value='1'> ��  <input name='Strip' type='radio' value='0' checked> ��</td>
					 </tr>
					 <tr class='tdbg'>
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eCommentID' value='1'></td>
							<td class='clefttitle' align='right'><strong>��������:</strong></td>
							<td><input name='Comment' type='radio' value='1'> ��  <input name='Comment' type='radio' value='0' checked> ��</td>
					</tr>
					 <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eSlide' value='1'></td>
							<td class='clefttitle' align='right'><strong>�Ƿ�õ�:</strong></td>
							<td><input name='Slide' type='radio' value='1'> ��  <input name='Slide' type='radio' value='0' checked> ��</td>
					 </tr>
					 <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eVerific' value='1'></td>
							<td class='clefttitle' align='right'><strong>�ĵ�״̬:</strong></td>
							<td><input name='verific' type='radio' value='1' checked> ����  <input name='Verific' type='radio' value='0'> δ��</td>
					</tr>
			    </table>
				
			<%if ChannelID=5 then%>		
			  </div>
			 </div>	
			<%end if%>	
				
				
			</TD>
		 </tr>
		 <tr class='tdbg'>
		    <td colspan=3 height='30'><b>˵����</b>��Ҫ�����޸�ĳ�����Ե�ֵ������ѡ�������ĸ�ѡ��Ȼ�����趨����ֵ��<br><div align='center'> <input type="submit" class="button" name="button1" value="ȷ������"> 
				  &nbsp; 
				  <%if ID<>"" then%>
				  <input type="reset" class="button" onClick="window.close()" name="button2" value=" �ر�ȡ�� ">
				  <%else%>
				  <input type="reset" class="button" name="button2" value=" ���� ">
				  <%end if%> </div></td>
		 </tr>

			  </form>
			</table>
			<br/>
			<br/>
			</div>
			</body>
			</html>
		 <%If Request.Form("Flag") = "Add" Then
		     If KS.G("choose")=0 Then
		      IDArr=Split(ID,",")
			 Else
			  IDArr=Split(Replace(KS.G("ClassID")," ",""),",")
			 End If
		      Set RS=Server.CreateObject("ADODB.RECORDSET")
			  For K=0 To Ubound(IDArr)
			  If KS.G("choose")=0 Then
			  RS.Open "Select * From " & KS.C_S(ChannelID,2) &" Where ID=" & IDArr(K), conn, 1, 3
			  Else
			  RS.Open "Select * From " & KS.C_S(ChannelID,2) &" Where Tid='" & IDArr(K) & "'", conn, 1, 3
			  End IF
			  If Not RS.EOF Then
			     Do While Not RS.Eof
				  If KS.ChkClng(KS.G("eTemplateID"))=1 And KS.G("TemplateID")<>"" Then RS("TemplateID") = KS.G("TemplateID")
				  If KS.ChkClng(KS.G("EWapTemplateID"))=1 And KS.G("WapTemplateID")<>"" Then RS("WapTemplateid")=KS.G("WapTemplateID")
				  If KS.ChkClng(KS.G("eKeyWords"))=1 Then
					 If InStr(" "&RS("KeyWords")&" "," "&KS.G("KeyWords")&" ") = 0 then
					  RS("KeyWords")  = RS("KeyWords")&" "&KS.G("KeyWords")
					 End If
				  End if
				  If KS.ChkClng(KS.G("eRank"))=1 And ChannelID<>8 Then       RS("Rank")      = KS.G("Rank")
				  If KS.ChkClng(KS.G("eAuthor"))=1 And ChannelID<>5 And ChannelID<>7 And ChannelID<>8 Then     RS("Author")    = KS.G("Author")
				  If KS.ChkClng(KS.G("eOrigin"))=1 And ChannelID<>5 And ChannelID<>7 And ChannelID<>8 Then     RS("Origin")    = KS.G("Origin")
				   Call SetAttributeField(RS)
				   
				   If ChannelID=5 Then '�̳����ü۸�
				       If KS.ChkClng(KS.G("ProductType"))<>0 Then 
						   RS("ProductType")      = KS.ChkClng(KS.G("ProductType"))
						   If KS.ChkClng(KS.G("ProductType"))=3 Then RS("Discount")=Request("DiscountPrice")
						   If KS.ChkClng(KS.G("EPriceMarket"))=1 Then   RS("Price_Market") = (RS("Price_Original")*(Request("DiscountPriceMarket")/10)*100)/100
						   If KS.ChkClng(KS.G("EPrice"))=1 Then   RS("Price") = (RS("Price_Original")*(Request("DiscountPrice")/10)*100)/100
						   If KS.ChkClng(KS.G("eScore"))=1 Then RS("Point")=KS.ChkClng((RS("Price_Original")*(Request("DiscountScore")/10)*100)/100)
						   If KS.ChkClng(KS.G("EPriceMember"))=1 Then 
							  If RS("GroupPrice")=0 Then
							   RS("Price_Member") = (RS("Price_Original")*(Request("DiscountPriceMember")/10)*100)/100
							  Else
								Dim RSG:Set RSG=Server.CreateObject("ADODB.RECORDSET")
								RSG.Open "Select * From KS_ProPrice Where ProID=" & RS("ID"),Conn,1,3
								Do While Not RSG.Eof
								  RSG("Price")=(RS("Price_Original")*(Request("price"&RSG("GroupID"))/10)*100)/100
								  RSG.Update
								  RSG.MoveNext
								Loop
								RSG.Close
								Set RSG=Nothing
							  End If
						   End If
					  End If
					  
					  If KS.ChkClng(KS.G("eLimitBuy"))<>0 Then
					     If KS.ChkCLng(KS.S("LimitBuyTaskID" & KS.ChkClng(KS.G("IsLimitbuy"))))=0 Then
						   KS.AlertHintScript "��ѡ������ID"
						   Response.End
						 End If
					     RS("IsLimitBuy")=KS.ChkClng(KS.G("IsLimitbuy"))
						 RS("LimitBuyPrice") = KS.S("LimitBuyPrice")
						 RS("LimitBuyAmount") = KS.ChkCLng(KS.S("LimitBuyAmount"))
						 RS("LimitBuyTaskID")=KS.ChkCLng(KS.S("LimitBuyTaskID" & KS.ChkClng(KS.G("IsLimitbuy"))))
						 
					   End If
					  
				   End If
				  
				   RS.Update
				 RS.MoveNext
				Loop
			 End If
			  RS.Close
			  
			  If KS.G("choose")=0 Then
			  RS.Open "Select * From [KS_ItemInfo] Where ChannelID=" & ChannelID &" And InfoID=" & IDArr(K), conn, 1, 3
			  Else
			  RS.Open "Select * From [KS_ItemInfo] Where Tid='" & IDArr(K) & "'", conn, 1, 3
			  End IF
			  If Not RS.EOF Then
			     Do While Not RS.Eof
				   Call SetAttributeField(RS)
				   RS.Update
				 RS.MoveNext
				Loop
			 End If
			  RS.Close
			  
			  
			 Next 
			 
			
			 
		   Set RS = Nothing
		   conn.Close:Set conn = Nothing
		   if ID<>"" then
		   KS.Echo "<script>alert('��ϲ���ɹ�������ѡ���ĵ�������!');window.close();</script>"
		   else
		   KS.Echo "<script>alert('��ϲ���������óɹ�!');</script>"
		   end if
		End If
		End Sub
		
		Sub SetAttributeField(RS)
				  If KS.ChkClng(KS.G("eHits"))=1 Then       RS("Hits")      =RS("Hits")+KS.ChkCLng(KS.G("Hits"))
				  If KS.ChkClng(KS.G("eRecommend"))=1 Then RS("Recommend") = KS.ChkCLng(KS.G("Recommend"))
				  If KS.ChkClng(KS.G("eRolls"))=1 Then     RS("Rolls")     = KS.ChkClng(KS.G("Rolls"))
				  If KS.ChkClng(KS.G("eStrip"))=1 Then     RS("Strip")     = KS.ChkClng(KS.G("Strip"))
				  If KS.ChkClng(KS.G("ePopular"))=1 Then   RS("Popular")   = KS.ChkClng(KS.G("Popular"))
				  If KS.ChkClng(KS.G("eCommentID"))=1 Then   RS("Comment")   = KS.ChkClng(KS.G("Comment"))
				  If KS.ChkClng(KS.G("eIsTop"))=1 Then     RS("IsTop")     = KS.ChkClng(KS.G("IsTop"))
				  If KS.ChkClng(KS.G("eSlide"))=1 Then     RS("Slide")     = KS.ChkCLng(KS.G("Slide"))
				  If KS.ChkClng(KS.G("eVerific"))=1 Then   RS("Verific")   = KS.ChkCLng(KS.G("Verific"))
				  If KS.ChkClng(KS.G("eAdddate"))=1 And IsDate(KS.G("AddDate")) Then  RS("AddDate")=KS.G("AddDate")
		End Sub

	
End Class
%> 
