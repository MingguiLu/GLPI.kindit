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
Set KSCls = New MyArticleCls
KSCls.Kesion()
Set KSCls = Nothing

Class MyArticleCls
        Private KS,KSUser,ChannelID,ID,ClassID,RS
		Private CurrentPage,totalPut,MaxPerPage
		Private ComeUrl,Selbutton,LoginTF,ReadPoint
		Private F_B_Arr,F_V_Arr,Title,FullTitle,KeyWords,Author,Origin,Intro,Content,Verific,PhotoUrl,Action,I,UserDefineFieldArr,UserDefineFieldValueStr,Province,City
		Private XmlFields,XmlFieldArr,Fi,IXml,INode
		Private Sub Class_Initialize()
			MaxPerPage =10
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
        
	   	
		Public Sub LoadMain()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=1
		LoginTF=Cbool(KSUser.UserLoginChecked)
		IF LoginTF=false  Then
		  Call KS.ShowTips("error","<li>�㻹û�е�¼���¼�ѹ��ڣ�������<a href='../user/login/'>��¼</a>!</li>")
		  Exit Sub
		End If
		If KS.C_S(ChannelID,6)<>1 Then Response.End()
		if KS.C_S(ChannelID,36)=0 then
		  Call KS.ShowTips("error","<li>��Ƶ��������Ͷ��!</li>")
		  Exit Sub
		end if
		
		'��������ͼ����
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
		F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")

		Call KSUser.Head()
		%>
		<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>">�ҷ�����<%=KS.C_S(ChannelID,3)%>(<span class="red"><%=Conn.Execute("Select count(id) from " & KS.C_S(ChannelID,2) &" where Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=1">�����(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=0">�����(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=2">�� ��(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="3" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=3">���˸�(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
        </div>
		<%
		Action=KS.S("Action")
		Select Case Action
		 Case "Del"	  Call KSUser.DelItemInfo(ChannelID,ComeUrl)
		 Case "Add","Edit"  Call DoAdd()
		 Case "DoSave" Call DoSave()
		 Case "refresh" Call KSUser.RefreshInfo(KS.C_S(ChannelID,2))
		 Case Else  Call ArticleList()
		End Select
	   End Sub
	   Sub ArticleList()
	      %>
			<script src="../ks_inc/jquery.imagePreview.1.0.js"></script>
		  <%		
		  
		    XmlFields=LFCls.GetConfigFromXML("usermodelfield","/modelfield/model",ChannelID)
			If Not KS.IsNul(XmlFields) Then
			 XmlFieldArr=Split(XmlFields,",")
			End If
            CurrentPage = KS.ChkClng(KS.S("page")): If CurrentPage<=0 Then CurrentPage=1
                                    
			Dim Param:Param=" Where Deltf=0 AND Inputer='"& KSUser.UserName &"'"
			Verific=KS.S("Status")
			If Verific="" or not isnumeric(Verific) Then Verific=4
            IF Verific<>4 Then Param= Param & " and Verific=" & Verific
			IF KS.S("Flag")<>"" Then
					  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
					  IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
			End if
			If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
			 Select Case Verific
				   Case 0 Call KSUser.InnerLocation("����" & KS.C_S(ChannelID,3) & "�б�")
				   Case 1 Call KSUser.InnerLocation("����" & KS.C_S(ChannelID,3) & "�б�")
				   Case 2 Call KSUser.InnerLocation("�ݸ�" & KS.C_S(ChannelID,3) & "�б�")
				   Case 3 Call KSUser.InnerLocation("�˸�" & KS.C_S(ChannelID,3) & "�б�")
                   Case Else Call KSUser.InnerLocation("����" & KS.C_S(ChannelID,3) & "�б�")
			 End Select
		 %>
								  <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="user_myarticle.asp?ChannelID=<%=ChannelID%>&Action=Add"><span style="font-size:14px;color:#ff3300">����<%=KS.C_S(ChannelID,3)%></span></a></div>

		<table  width="99%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
         <%
		     Dim FieldStr:FieldStr="ID,Tid,Title,Inputer,AddDate,PhotoUrl,Verific,Recommend,Popular,Strip,Rolls,Slide,IsTop,Hits,IsVideo"
			 If IsArray(XmlFieldArr) Then
			 For Fi=0 To Ubound(XmlFieldArr)
			  if lcase(Split(XmlFieldArr(fi),"|")(1))<>"modeltype" and lcase(Split(XmlFieldArr(fi),"|")(1))<>"attribute" and ks.foundinarr(lcase(FieldStr),lcase(Split(XmlFieldArr(fi),"|")(1)),",")=false then
			   FieldStr=FieldStr & "," & Split(XmlFieldArr(fi),"|")(1)
			  end if
			 Next
			End If
			Dim Sql:sql = "select " & FieldStr & " from " & KS.C_S(ChannelID,2) & Param &" order by ID Desc"
			 Set RS=Server.CreateObject("AdodB.Recordset")
			  RS.open sql,conn,1,1
			  If RS.EOF And RS.BOF Then
			   RS.Close : Set RS=Nothing
			  Response.Write "<tr><td class='tdbg' align='center' colspan=12 height=30 valign=top>��ǰû���κ�" & KS.C_S(ChannelID,3) & "!</td></tr>"
			 Else
				totalPut = RS.RecordCount
				If CurrentPage < 1 Then	CurrentPage = 1
					
				If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrentPage - 1) * MaxPerPage
				Else
					CurrentPage = 1
				End If
				Set IXML=KS.ArrayToxml(RS.GetRows(MaxPerPage),rs,"row","")
				RS.Close : Set RS=Nothing
				If IsArray(XmlFieldArr) Then
				 Call ShowDiyList
				Else
				 Call showContent
				End If
			End If
     %>
	  </table>
	  
			 <table cellspacing="0" cellpadding="0" border="0" width="100%">
				 <tr>
				 <td>
								 <label><input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;ѡ������</label>&nbsp;<button id="btn1"  class="pn pnc" onClick="return(confirm('ȷ��ɾ��ѡ�е�<%=KS.C_S(ChannelID,3)%>��?'));" type=submit><strong>ɾ��ѡ��</strong></button></FORM>       
				 </td>
				 <td align='right'>
									 <%
							         Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
						    	      %>
				 </td>
				 </tr>
			 </table>
									 
	<table>				
	 
	 <tr class='tdbg'>
           <form action="User_MyArticle.asp?ChannelID=<%=ChannelID%>" method="post" name="searchform">
           <td height="45" colspan=14>
				<strong><%=KS.C_S(ChannelID,3)%>������</strong>
				 <select name="Flag">
					<option value="0">����</option>
					<option value="1">�ؼ���</option>
				 </select>
										  
				�ؼ���
				<input type="text" name="KeyWord" class="textbox" onclick="if(this.value=='�ؼ���'){this.value=''}" value="�ؼ���" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" �� �� ">
			 </td>
			 </form>
             </tr>
         </table>
	</div>
 <%
  End Sub
  
  Sub ShowDiyList()
  %>
  <tr align="center" class="title">
   <td><b>ѡ��</b></td><td><b>����</b></td>
   <%
   If IsArray(XmlFieldArr) Then
	 For Fi=0 To Ubound(XmlFieldArr)
	   KS.echo ("<td nowrap>" & Split(XmlFieldArr(fi),"|")(0) & "</td>")
	 Next
   End If
   %>
   <td><b>����</b></td>
  </tr>
  <%
   For Each INode In IXml.DocumentElement.SelectNodes("row")
    Dim AttributeStr:AttributeStr = ""
	If Instr(lcase(XmlFields),"attribute")<>0 then
		If Cint(INode.SelectSingleNode("@recommend").text) = 1 Or Cint(INode.SelectSingleNode("@popular").text) = 1 Or Cint(INode.SelectSingleNode("@strip").text) = 1 Or Cint(INode.SelectSingleNode("@rolls").text) = 1 Or Cint(INode.SelectSingleNode("@slide").text) = 1 Or Cint(INode.SelectSingleNode("@istop").text) = 1 Then
			If Cint(INode.SelectSingleNode("@recommend").text) = 1 Then AttributeStr = AttributeStr & (" <span title=""�Ƽ�" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""green"">��</font></span>&nbsp;")
			If Cint(INode.SelectSingleNode("@popular").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""����" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""red"">��</font></span>&nbsp;")
			If Cint(INode.SelectSingleNode("@strip").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""����ͷ��"" style=""cursor:default""><font color=""#0000ff"">ͷ</font></span>&nbsp;")
			If Cint(INode.SelectSingleNode("@rolls").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""����" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""#F709F7"">��</font></span>&nbsp;")
			If Cint(INode.SelectSingleNode("@slide").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""�õ�Ƭ" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""black"">��</font></span>")
			IF Cint(INode.SelectSingleNode("@istop").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""�̶�" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""brown"">��</font></span>")
			If KS.C_S(Channelid,6)=1 Then
			IF KS.ChkClng(INode.SelectSingleNode("@isvideo").text) = 1 Then AttributeStr = AttributeStr & ("<span title=""��Ƶ" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""#ff6600"">Ƶ</font></span>")
			End If
			If AttributeStr="" Then AttributeStr="---"
		Else
			AttributeStr = "---"
		End If
	End If
  %>
   <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td class="splittd" align="center"><input id="ID" type="checkbox" value="<%=INode.SelectSingleNode("@id").text%>"  name="ID"></td>
	<td class="splittd"><a href="../item/show.asp?m=<%=ChannelID%>&d=<%=INode.SelectSingleNode("@id").text%>" target="_blank"><%=KS.Gottopic(INode.SelectSingleNode("@title").text,30)%></a></td>
	<%
	If IsArray(XmlFieldArr) Then
		For Fi=0 To Ubound(XmlFieldArr)
			KS.echo ("<td class='splittd' nowrap align='center'>&nbsp;")
		   select case lcase(Split(XmlFieldArr(fi),"|")(1))
				    case "modeltype" KS.echo KS.C_S(ChannelID,3)
					case "attribute" KS.echo AttributeStr
					case "adddate" ks.echo KS.GetTimeFormat(INode.SelectSingleNode("@adddate").text)
					case "refreshtf" 
						If KS.C_S(ChannelId,7)="0" then
						  ks.echo "<span style='color:blue;cursor:default' title='��ģ��û���������ɾ�̬HTML,��������'>��������</span>"
					   Else
						   if INode.SelectSingleNode("@refreshtf").text="1" then
								     ks.echo "<font color=green>������</font>"
						   else 
								     ks.echo "<font color='#ff3300'>δ����</font>"
						   end if
					   End If
					case else
					  ks.echo INode.SelectSingleNode("@" &lcase(Split(XmlFieldArr(fi),"|")(1))).text
					end  select
			ks.echo ("&nbsp;</td>")
	 Next
	End If
	%>
	<td class="splittd" align="center">
	<%If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(3))=1 Then%>
		<a href="?ChannelID=<%=ChannelID%>&action=refresh&id=<%=INode.SelectSingleNode("@id").text%>" class="box">ˢ��</a>
	<%end if%>
	<%if cint(INode.SelectSingleNode("@verific").text)<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a class='box' href="User_MyArticle.asp?channelid=<%=channelid%>&id=<%=INode.SelectSingleNode("@id").text%>&Action=Edit&&page=<%=CurrentPage%>">�޸�</a> <a class='box' href="User_MyArticle.asp?channelid=<%=channelid%>&action=Del&ID=<%=INode.SelectSingleNode("@id").text%>" onclick = "return (confirm('ȷ��ɾ��<%=KS.C_S(ChannelID,3)%>��?'))">ɾ��</a>
	<%else
		If KS.C_S(ChannelID,42)=0 Then
			Response.write "---"
		Else
			Response.Write "<a  class='box' href='?channelid=" & channelid & "&id=" & INode.SelectSingleNode("@id").text &"&Action=Edit&&page=" & CurrentPage &"'>�޸�</a> <a class='box' href='#' disabled>ɾ��</a>"
		End If
	end if
	%>
	</td>
   </tr>
  <%
   Next
  End Sub
  
  Sub ShowContent()
    Dim I,PhotoUrl
    Response.Write "<FORM Action=""User_MyArticle.asp?ChannelID=" & ChannelID & "&Action=Del"" name=""myform"" method=""post"">"
    For Each INode In IXml.DocumentElement.SelectNodes("row")
        If Not KS.IsNul(INode.SelectSingleNode("@photourl").text) Then
		 PhotoUrl=INode.SelectSingleNode("@photourl").text
		Else
		 PhotoUrl="Images/nopic.gif"
		End If %>
           <tr>
			<td class="splittd" width="10"><input id="ID" type="checkbox" value="<%=INode.SelectSingleNode("@id").text%>"  name="ID"></td>
		    <td class="splittd" width="33"><div style="cursor:pointer;text-align:center;width:33px;height:33px;border:1px solid #f1f1f1;padding:1px;"><a href="<%=PhotoUrl%>" target="_blank" title="<%=INode.SelectSingleNode("@title").text%>" class="preview"><img  src="<%=PhotoUrl%>" width="32" height="32"></a></div>
			</td>
            <td height="45" align="left" class="splittd">
						<div class="ContentTitle"><a href="../item/show.asp?m=<%=ChannelID%>&d=<%=INode.SelectSingleNode("@id").text%>" target="_blank"><%=trim(INode.SelectSingleNode("@title").text)%></a>
						</div>
						
						<div class="Contenttips">
			            <span>
						 ��Ŀ��[<%=KS.C_C(INode.SelectSingleNode("@tid").text,1)%>] �����ˣ�<%=INode.SelectSingleNode("@inputer").text%> ����ʱ�䣺<%=KS.GetTimeFormat(INode.SelectSingleNode("@adddate").text)%>
						 ״̬��<%Select Case cint(INode.SelectSingleNode("@verific").text)
									Case 0
									   Response.Write "<span style=""color:green"">����</span>"
									Case 1
									   Response.Write "<span>����</span>"
                                    Case 2
									  Response.Write "<span style=""color:red"">�ݸ�</span>"
									 Case 3
									  Response.Write "<span style=""color:blue"">�˸�</span>"
                               end select
							 %>
						 </span>
						</div>
						</td>
                        <td class="splittd" align="center">
						<%If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(3))=1 Then%>
						   <a href="?ChannelID=<%=ChannelID%>&action=refresh&id=<%=INode.SelectSingleNode("@id").text%>" class="box">ˢ��</a>
						<%end if%>
							<%if cint(INode.SelectSingleNode("@verific").text)<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a class='box' href="User_MyArticle.asp?channelid=<%=channelid%>&id=<%=INode.SelectSingleNode("@id").text%>&Action=Edit&&page=<%=CurrentPage%>">�޸�</a> <a class='box' href="User_MyArticle.asp?channelid=<%=channelid%>&action=Del&ID=<%=INode.SelectSingleNode("@id").text%>" onclick = "return (confirm('ȷ��ɾ��<%=KS.C_S(ChannelID,3)%>��?'))">ɾ��</a>
							<%else
								  If KS.C_S(ChannelID,42)=0 Then
									  Response.write "---"
								  Else
									  Response.Write "<a  class='box' href='?channelid=" & channelid & "&id=" & INode.SelectSingleNode("@id").text &"&Action=Edit&&page=" & CurrentPage &"'>�޸�</a> <a class='box' href='#' disabled>ɾ��</a>"
								  End If
							end if
							%>
						</td>
                       </tr>
 <%
   Next

 End Sub

%>
<!--#include file="../ks_cls/UserFunction.asp"-->
<%
 '�������
 Sub DoAdd()
        ID=KS.ChkClng(KS.S("id"))
        Session("UploadFileIDs")=""  '���渽��ID��
        Call KSUser.InnerLocation("����"& KS.C_S(ChannelID,3))
		If ID<>0 Then
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
	     RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName &"' and ID=" & ID,Conn,1,1
		 If Not RS.Eof Then
		  ClassID=RS("Tid") : SelButton=KS.C_C(ClassID,1)
		 End If
		Else
		 SelButton="ѡ����Ŀ..."
	    End If
		%>
		<script type="text/javascript" src="../editor/ckeditor.js"></script>
		<script type="text/javascript" src="../ks_inc/kesion.box.js"></script>
		<script language = "JavaScript">
		   function GetKeyTags(){
			  var text=escape($('#Title').val());
			  if (text!=''){
				  $('#KeyWords').val('���Ե�,ϵͳ�����Զ���ȡtags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data)).attr("disabled",false);
				  });
			  }else{
			   alert('�Բ���,�����������!');
			  }
			}
		    function CheckClassID(){
				if (document.myform.ClassID.value=="0" || document.myform.ClassID.value=='') {
					alert("��ѡ��<%=KS.C_S(ChannelID,3)%>��Ŀ��");
					return false;}		
				  return true;
			}
			function insertHTMLToEditor(codeStr){ CKEDITOR.instances.Content.insertHtml(codeStr);} 
			function CheckForm(){
				<%Call KSUser.ShowUserFieldCheck(ChannelID)%>
				if (document.myform.ClassID.value=="0") 
				  {
					alert("��ѡ��<%=KS.C_S(ChannelID,3)%>��Ŀ��");
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>���⣡");
					document.myform.Title.focus();
					return false;
				  }	
				<%if F_B_Arr(9)=1 Then%> 
				    if (CKEDITOR.instances.Content.getData()=="")
					{
					  alert("<%=KS.C_S(ChannelID,3)%>���ݲ������գ�");
					  CKEDITOR.instances.Content.focus();
					  return false;
					}
				<%end if%>
				 return true; }
		</script>
		<form  action="User_MyArticle.asp?channelid=<%=channelid%>&Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
		<%
		' GetInputForm false,ChannelID,KS.ChkClng(KS.S("id")),KSUser,rs
		' ks.die ""
		Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
		If KS.IsNul(XmlForm) Then 
		 GetInputForm false,ChannelID,KS.ChkClng(KS.S("id")),KSUser,rs
		Else
		   UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
		   If Action="Edit" Then
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
			  If UserDefineFieldValueStr<>"0" And UserDefineFieldValueStr<>""  Then UserDefineFieldValueStr=Split(UserDefineFieldValueStr,"||||")
		  End If
		 Scan XmlForm
		'  ks.echo XmlForm
		End If
		%>
		</form>
		<%
  End Sub
  
 
  
 Sub DoSave()
    ClassID=KS.S("ClassID")
	ID=KS.ChkClng(KS.S("ID"))
	If KS.ChkClng(KS.C_C(ClassID,20))=0 Then
	 Response.Write "<script>alert('�Բ���,ϵͳ�趨�����ڴ���Ŀ����,��ѡ��������Ŀ!');history.back();</script>":Exit Sub
	End IF
	Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
	KeyWords=KS.LoseHtml(KS.S("KeyWords"))
	Author=KS.LoseHtml(KS.S("Author"))
	Origin=KS.LoseHtml(KS.S("Origin"))
	Content = Request.Form("Content")
	Content=KS.FilterIllegalChar(KS.ClearBadChr(content))
				 
	if KS.IsNul(Content) Then Content="&nbsp;"
	Verific=KS.ChkClng(KS.S("Status"))
	Intro  = KS.FilterIllegalChar(KS.LoseHtml(KS.S("Intro")))
	Province= KS.LoseHtml(KS.S("Province"))
	City    = KS.LoseHtml(KS.S("City"))
	FullTitle = KS.LoseHtml(KS.S("FullTitle"))
	if Intro="" And KS.ChkClng(KS.S("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(Request.Form("Content")),200)
				 
	Dim Fname,FnameType,TemplateID,WapTemplateID
	If KS.ChkClng(KS.S("ID"))=0 Then
					 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
					 RSC.Open "select top 1 * from KS_Class Where ID='" & ClassID & "'",conn,1,1
					 if RSC.Eof Then 
					  Response.end
					 Else
					 FnameType=RSC("FnameType")
					 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
					 TemplateID=RSC("TemplateID")
					 WapTemplateID=RSC("WapTemplateID")
					 End If
					 RSC.Close:Set RSC=Nothing
	End If
	If KS.ChkClng(KS.C_S(ChannelID,17))<>0 And Verific=0 Then Verific=1
	If ID<>0 and verific=1  Then
		If KS.ChkClng(KS.C_S(ChannelID,42))=2 Then Verific=1 Else Verific=0
	End If
	if KS.C_S(ChannelID,42)=2 and KS.ChkClng(KS.S("okverific"))=1 Then verific=1
	If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '����VIP�û��������
				 
	PhotoUrl=KS.S("PhotoUrl")
	Call KSUser.CheckDiyField(ChannelID,UserDefineFieldArr)
				
	If ClassID="" Then
		KS.Die "<script>alert('��û��ѡ��" & KS.C_S(ChannelID,3) & "��Ŀ!');history.back();</script>"
	 End IF
	If Title="" Then
		KS.Die "<script>alert('��û������" & KS.C_S(ChannelID,3) & "����!');history.back();</script>"
	End IF
	If Content="" and KS.ChkClng(F_B_Arr(9))=1 Then
		KS.Die "<script>alert('��û������" & KS.C_S(ChannelID,3) & "����!');history.back();</script>"
	End IF
	Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
	RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName & "' and ID=" & ID,Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("Adddate")=Now
				  RSObj("Rank")="����"
				  RSObj("Inputer")=KSUser.UserName
				 End If
				  RSObj("Title")=Title
				  RSObj("FullTitle")=FullTitle
				  RSObj("Tid")=ClassID
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Origin")=Origin
				  RSObj("ArticleContent")=Content
				  RSObj("Verific")=Verific
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("Intro")=Intro
				  RSObj("DelTF")=0
				  RSObj("Comment")=1
                  If F_B_Arr(18)=1 Then
				  RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
				  RSObj("Province")=Province
				  RSObj("City")=City				  
				  if PhotoUrl<>"" Then 
				   RSObj("PicNews")=1
				  Else
				   RSObj("PicNews")=0
				  End if
				  If F_B_Arr(25)="1" Then	RSObj("MapMarker")=KS.S("MapMark")
				  Call KSUser.AddDiyFieldValue(RSObj,UserDefineFieldArr)
				  
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" And KS.ChkClng(KS.S("ID"))=0 Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				Fname=RSOBj("Fname")
				
				If Verific=1 Then 
				    Call KS.SignUserInfoOK(ChannelID,KSUser.UserName,Title,InfoID)
					If KS.C_S(ChannelID,17)=2  and (KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2) Then
					 Dim KSRObj:Set KSRObj=New Refresh
					 Dim DocXML:Set DocXML=KS.RsToXml(RSObj,"row","root")
				     Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
					  KSRObj.ModelID=ChannelID
					  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
					  Call KSRObj.RefreshContent()
					  Set KSRobj=Nothing
					End If
				End If
				 RSObj.Close:Set RSObj=Nothing
				 
				If Not KS.IsNul(Session("UploadFileIDs")) Then 
				 Conn.Execute("Update [KS_UpLoadFiles] Set InfoID=" & InfoID &",classID=" & KS.C_C(ClassID,9) & " Where ID In (" & KS.FilterIds(Session("UploadFileIDs")) & ")")
				End If

				 
               If ID=0 Then
			     Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Intro,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
                 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl ,0)
			     Call KSUser.AddLog(KSUser.UserName,"����Ŀ[<a href='" & KS.GetFolderPath(ClassID) & "' target='_blank'>" & KS.C_C(ClassID,1) & "</a>]������" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""!",1)
				 KS.Echo "<script>if (confirm('" & KS.C_S(ChannelID,3) & "��ӳɹ������������?')){location.href='User_myArticle.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';}else{location.href='User_MyArticle.asp?ChannelID=" & ChannelID & "';}</script>"
			   Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Intro,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl ,1)
			     Call KSUser.AddLog(KSUser.UserName,"��" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""�����޸�!",1)
				 KS.Echo "<script>alert('" & KS.C_S(ChannelID,3) & "�޸ĳɹ�!');location.href='User_MyArticle.asp?channelid=" & channelid & "';</script>"
			   End If
  End Sub
  
End Class
%> 
