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
Set KSCls = New User_Photo
KSCls.Kesion()
Set KSCls = Nothing

Class User_Photo
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,AddDate,Weather,PhotoUrls,descript
		Private XCID,Title,Tags,UserName,Face,Content,Status,PicUrl,Action,I,ClassID,password
		Private Sub Class_Initialize()
		  MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
       Public Sub loadMain()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		ElseIf KS.SSetting(0)=0 Then
		 Call KS.Alert("�Բ��𣬱�վ�رո��˿ռ书�ܣ�","")
		 Exit Sub
		ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)=0 Then
		 Call KS.Alert("�㲻�ԣ��㻹û�п�ͨ�ռ书�ܣ�","User_Blog.asp")
		 Exit Sub
		ElseIf Conn.Execute("Select status From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)<>1 Then
		    Response.Write "<script>alert('�Բ�����Ŀռ仹û��ͨ����˻�������');history.back();</script>"
			response.end
		End If

		Call KSUser.Head()
		Call KSUser.InnerLocation("�ҵ����")
		KSUser.CheckPowerAndDie("s05")
		%>
		<div class="tabs">	
		   <ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="?">�ҵ����</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="?Status=1">�������(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=1")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="?Status=0">�������(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=0")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="?Status=2">�������(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=2")(0)%></span>)</a></li>
			</ul>
        </div>
			 <div style="padding-left:20px;"><img src="images/fav.gif" align="absmiddle"><a href="User_Photo.asp?Action=Add"><span style="font-size:14px;color:#ff3300">�ϴ���Ƭ</span></a>
			 
			 </div>

		<%

			Select Case KS.S("Action")
			 Case "Del"
			  Call Delxc()
			 Case "Delzp"
			  Call Delzp()
			 Case "Editzp"
			  Call Editzp()
			 Case "Add"
			  Call Addzp()
			 Case "AddSave"
			  Call AddSave()
			 Case "EditSave"
			  Call EditSave()
			 Case "ViewZP"
			  Call ViewZP()
			 Case "Editxc","Createxc"
			  Call Managexc()
			 Case "photoxcsave"
			  Call photoxcsave()
			 Case Else
			  Call PhotoxcList()
			End Select
	   End Sub
	   '�鿴��Ƭ
	   Sub ViewZP()
	    Dim title
	    Dim xcid:xcid=KS.Chkclng(KS.S("XCID"))
	    Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "select top 1 xcname from KS_Photoxc WHERE ID=" & XCID,CONN,1,1
		if rs.Eof And RS.Bof Then 
		 rs.close:set rs=nothing
		 response.write "<script>alert('�������ݳ���');history.back();</script>"
		 response.end
		end if
		title=rs(0)
		rs.close
		Call KSUser.InnerLocation("�鿴��Ƭ")
	  			  %>
			   
	   		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
            <tr class="title">
              <td align=center colspan=5><%=Title%></td>
            </tr>
			<%
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
			 rs.open "select * from KS_PhotoZP where xcid=" & xcid,conn,1,1
			if rs.eof and rs.bof then
			  response.write "<tr class='tdbg'><td  height='30' colspan='5'>�������û����Ƭ����<a href=""?action=Add&xcid=" & xcid &""">�ϴ�</a>��</td></tr>"
			else
			 				  MaxPerPage =5
								totalPut = RS.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call showzplist(xcid)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call showzplist(xcid)
									Else
										CurrentPage = 1
										Call showzplist(xcid)
									End If
								End If
        end if%>
      </table>
	  <div style="padding-right:30px">
	  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
	  </div>
<%End Sub
sub showzplist(xcid)
%>
    <script type="text/javascript" src="../ks_inc/highslide/highslide.js"></script>
    <link href="../ks_inc/highslide/highslide.css" type=text/css rel=stylesheet>
	<script type="text/javascript">
		hs.graphicsDir = '/ks_inc/highslide/graphics/';
		hs.transitions = ['expand', 'crossfade'];
		hs.wrapperClassName = 'dark borderless floating-caption';
		hs.fadeInOut = true;
		hs.dimmingOpacity = .75;
		
		if (hs.addSlideshow) hs.addSlideshow({
			interval: 5000,
			repeat: false,
			useControls: true,
			fixedControls: 'fit',
			overlayOptions: {
				opacity: .6,
				position: 'bottom center',
				hideOnMouseOut: true
			}
		});
	</script>
<%
     Dim I
    Response.Write "<FORM Action=""?Action=Delzp"" name=""myform"" method=""post"">"
			 do while not rs.eof
			 %>
			<tr class="tdbg"> 
            
          </tr>
          <tr class="tdbg"> 
<td width="16%" rowspan="4">
			<table border="0" align="center" cellpadding="2" cellspacing="1" class="border">
                <tr> 
                  <td><a href="<%=rs("photourl")%>" class="highslide" onClick="return hs.expand(this)"  title="<%=rs("title")%>"><img src="<%=rs("photourl")%>" width="85" height="100" border="0"></a>
                  </td>
                </tr>
              </table></td>		  
            <td><div align="center">�������ڣ�</div></td>
            <td><%=rs("adddate")%></td>
            <td><div align="center">ͼƬ��С��</div></td>
            <td><%=rs("photosize")%>byte   </td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">��Ƭ��ַ��</div></td>
            <td colspan="3"><%=rs("photourl")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">��Ƭ������</div></td>
            <td colspan="3"><%=rs("descript")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">���������</div></td>
            <td><%=rs("hits")%> �˴�</td>
            <td colspan="2" height="28"><div align="center"><a href="?Action=Editzp&Id=<%=rs("id")%>" class="box">�޸�</a> <a href="?id=<%=rs("id")%>&Action=Delzp" onClick="{if(confirm('ȷ��ɾ������Ƭ��')){return true;}return false;}" class="box">ɾ��</a> 
                <INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
              </div></td>
          </tr>
          <tr> 
            <td colspan="5" height="3" class="splittd">&nbsp;</td>
          </tr>
			<% rs.movenext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
			 loop
		 %>
		 <tr class="tdbg">
		   <td colspan="5" align="right">
		  								&nbsp;&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;ѡ�б�ҳ��ʾ��������Ƭ&nbsp;<INPUT class="button" onClick="return(confirm('ȷ��ɾ��ѡ�е���Ƭ��?'));" type=submit value=ɾ��ѡ������Ƭ name=submit1>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        </td>
		 </tr>
		 </form>
		 <%
	   End Sub
	    '��ᣬ��ӣ��޸�
	   Sub Managexc()
	    Dim xcname,ClassID,Descript,PhotoUrl,PassWord,ListReplayNum,ListGuestNum,OpStr,TipStr,TemplateID,Flag,ListLogNum
		Dim ID:ID=KS.ChkCLng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_Photoxc Where ID=" & ID,conn,1,1
		If Not RS.EOF Then
		Call KSUser.InnerLocation("�޸����")
		 xcname=RS("xcname")
		 ClassID=RS("ClassID")
		 Descript=RS("Descript")
		 flag=RS("Flag")
		 PhotoUrl=RS("PhotoUrl")
		 PassWord=RS("PassWord")
		 OpStr="OK�ˣ�ȷ���޸�":TipStr="�� �� �� �� �� ��"
		Else
		 Call KSUser.InnerLocation("�������")
		 xcname=FormatDatetime(Now,2)
		 ClassID="0"
		 flag="1"
		 PhotoUrl=""
		 OpStr="OK�ˣ���������":TipStr="�� �� �� �� �� ��"
		End if
		RS.Close:Set RS=Nothing
	    %>
		<script>
		 function CheckForm()
		 {
		  if (document.myform.xcname.value=='')
		  {
		   alert('�������������!');
		   document.myform.xcname.focus();
		   return false;
		  }
		  if (document.myform.ClassID.value=='0')
		  {
		   alert('��ѡ���������!');
		   document.myform.ClassID.focus();
		   return false;
		  }
		  return true;
		 }

		</script>
		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
          <form  action="User_Photo.asp?Action=photoxcsave&id=<%=id%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
            <tr class="title">
              <td colspan=2 align=center><%=TipStr%></td>
            </tr>
            <tr class="tdbg">
              <td  height="25" class="clefttitle">������ƣ�</td>
              <td><input class="textbox" name="xcname" type="text" id="xcname" style="width:230px; " value="<%=xcname%>" maxlength="100" />
              <span style="color: #FF0000">*</span><span class="msgtips">���������ȡ�����ʵ�����,�����д�漯��</span></td>
            </tr>
<tr class="tdbg">
              <td class="clefttitle" height="25">�����ࣺ</td>
              <td><select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-��ѡ�����-</option>
                    <% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_PhotoClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If ClassID=RS("ClassID") Then
								  Response.Write "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                  </select>          <span class="msgtips">�����࣬�Ա�������</span>     </td>
            </tr>
			<tr class="tdbg"> 
                  <td class="clefttitle">�Ƿ񹫿���</td>
                  <td><select style="width:160px" onChange="if(this.options[selectedIndex].value=='3'){document.myform.all.mmtt.style.display='block';}else{document.myform.all.mmtt.style.display='none';}"  name="flag">
                      <option value="1"<%if flag="1" then response.write " selected"%>>��ȫ����</option>
                      <option value="2"<%if flag="2" then response.write " selected"%>>��Ա����</option>
                      <option value="3"<%if flag="3" then response.write " selected"%>>���빲��</option>
                      <option value="4"<%if flag="4" then response.write " selected"%>>��˽���</option>
                    </select><span class="msgtips">��������Ϊֻ��Ȩ�޵��û���������� </span><span class=child id=mmtt name="mmtt" <%if flag<>3 then%>style="display:none;"<%end if%>>���룺<input type="password" name="password" style="width:160px" maxlength="16" value="<%=password%>" size="20"></span> 
				   </td>
            </tr>
            <tr class="tdbg">
              <td class="clefttitle">�����棺</td>
              <td><input class="textbox" name="PhotoUrl" type="text" id="PhotoUrl" style="width:230px; " value="<%=PhotoUrl%>" />                  <span class="msgtips">ֻ֧��jpg��gif��png��С��50k��Ĭ�ϳߴ�Ϊ85*100</span>
				  <div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=9998' frameborder="0" align="center" width='100%' height='30' scrolling="no"></iframe>
				  </div>
				  </td>
            </tr>
            <tr class="tdbg">
              <td class="cleftittle">�����ܣ� </td>
              <td><textarea class="textarea" name="Descript" id="Descript" cols=50 rows=6><%=Descript%></textarea>              <span class="msgtips">���ڴ����ļ�Ҫ����˵����</span>
				  </td>
            </tr>
            <tr class="tdbg">
			  <td></td>
              <td>
			    <button class="pn" type="submit"><strong><%=OpStr%></strong></button>
              </td>
            </tr>
          </form>
</table>
		<%
	   End Sub
	   '�������
	   Sub photoxcsave()
	     Dim xcname:xcname=KS.S("xcname")
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Descript:Descript=KS.S("Descript")
		 Dim Flag:Flag=KS.S("Flag")
		 Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
		 Dim PassWord:PassWord=KS.S("PassWord")
		 Dim ID:ID=KS.Chkclng(KS.S("id"))
		 If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="/images/user/nopic.gif"
		 If xcname="" Then Response.Write "<script>alert('�������������!');history.back();</script>"
		 If ClassID=0 Then Response.Write "<script>alert('��ѡ���������!');history.back();</script>"
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Photoxc Where id=" & id ,conn,1,3
		 If RS.Eof And RS.Bof Then
		   RS.AddNew
		    RS("AddDate")=now
			if ks.SSetting(4)=1 then
			RS("Status")=0 '��Ϊ����
			else
			RS("Status")=1 '��Ϊ����
			end if
		 End If
		    RS("UserName")=KSUser.UserName
		    RS("xcname")=xcname
			RS("ClassID")=ClassID
			RS("Descript")=Descript
			RS("Flag")=Flag
			RS("Password")=PassWord
			RS("PhotoUrl")=PhotoUrl
		  RS.Update
		  RS.MoveLast
		  ID=rs("id")
		  RS.Close:Set RS=Nothing
		  If KS.Chkclng(KS.S("id"))=0 Then
		   Call KS.FileAssociation(1028,ID,PhotoUrl,0)
		   Call KSUser.AddLog(KSUser.UserName,"���������!����: "&xcname & " <a href=""../space/?" & KSUser.UserName & "/showalbum/" & id & """ target=""_blank"">�鿴</a>",104)
		   Response.Write "<script>alert('��ϲ!��ᴴ���ɹ�,�����ϴ���Ƭ');location.href='User_Photo.asp?action=Add&xcid=" & id &"';</script>"
		  Else
		   Call KS.FileAssociation(1028,ID,PhotoUrl,1)
		   Call KSUser.AddLog(KSUser.UserName,"�޸������!����: "&xcname & " <a href=""../space/?" & KSUser.UserName & "/showalbum/" & id & """  target=""_blank"">�鿴</a>",104)
		   Response.Write "<script>alert('����޸ĳɹ�!');location.href='User_Photo.asp';</script>"
		  End If
	   End Sub


	  
	   '����б�
	   Sub PhotoxcList()
			  
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
									IF KS.S("status")<>"" Then
									  Param=Param & " And status=" & KS.ChkClng(KS.S("status"))
									End if
									
									
									'If KS.S("XCID")<>"" And KS.S("XCID")<>"0" Then Param=Param & " And XCID=" & KS.ChkClng(KS.S("XCID")) & ""
									Dim Sql:sql = "select * from KS_Photoxc "& Param &" order by AddDate DESC"


								    Call KSUser.InnerLocation("��������б�")
								  %>
								     
				                     <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1">
                                                <tr class="title">
                                                  <td colspan="6" height="22" align="center">�� �� �� ��</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>����û�д������!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
											If CurrentPage < 1 Then
												CurrentPage = 1
											End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowXC
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowXC
									Else
										CurrentPage = 1
										Call ShowXC
									End If
								End If
				End If
     %>                      
                        </table>
		  <%
  End Sub
  
  Sub ShowXC()
     Dim I,K
   Do While Not RS.Eof
         %>
           <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		   <%
		   For K=1 To 4
		   %>
            <td width="25%" height="22" align="center">
									  <table width=154 height=185 border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF" id=AutoNumber2 style="BORDER-COLLAPSE: collapse">
										  <td width=123 height=185>
											<table id=AutoNumber3 style="BORDER-COLLAPSE: collapse" borderColor=#b2b2b2 height=179 cellSpacing=0 cellPadding=0 width="117%" border=0>
											  <tr>
												<td width="100%" height=179>
												  <table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width="99%" border=0>
													<tr>
													  <td align=middle width="100%" height=22><B><a href="?xcid=<%=rs("id")%>&action=ViewZP"><%=ks.gottopic(rs("xcname"),18)%></a></B><%select case rs("status")
													     case 1:response.write "[����]"
														 case 2:response.write "<font color=blue>[����]</font>"
														 case 0:response.write "<font color=red>[δ��]</font>"
														end select
														%>
													  </td>
													</tr>
													<tr>
													  <td align=middle width="100%">
														<table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0>
														  <tr>
															<td background="images/pic.gif" width="136" height="106" valign="top"><a href="?xcid=<%=rs("id")%>&action=ViewZP"><img style="margin-left:6px;margin-top:5px" src="<%=rs("photourl")%>" width="120" height="90" border=0></a></td>
														  </tr>
														</table>
													  </td>
													</tr>
													<tr>
													  <td align=middle width="100%" height=23><%=rs("xps")%>��/<%=rs("hits")%>��</td>
													</tr>
													<tr>
													  <td align=middle width="100%" height=23><a href="?Action=Editxc&id=<%=rs("id")%>">�޸�</a>&nbsp;<a href="?Action=Del&id=<%=rs("id")%>" onClick="return(confirm('ɾ����Ὣɾ����������������Ƭ��ȷ��ɾ����'))">ɾ��</a>&nbsp;
													  <% select case rs("flag")
													      case 1
													       response.write "<font color=red>[����]</font>"
														  case 2
													       response.write "<font color=red>[��Ա]</font>"
														  case 3
													       response.write "<font color=red>[����]</font>"
														  case 4
													       response.write "<font color=red>[��˽]</font>"
														 end select
													%>
													  </td>
													</tr>
												  </table>
												</td>
											  </tr>
											</table>
										  </td>
										</tr>
			  </table>
			 </td>
                       
					                  <%
							RS.MoveNext
							I=I+1
					  If I >= MaxPerPage Or RS.Eof Then Exit For
				  Next
			      do While K<4 
				   response.write "<td width=""25%""></td>"
				   k=k+1
				  Loop%>
		    </tr>
				 <%
					  If I >= MaxPerPage Or RS.Eof Then Exit do
	   Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=6 valign=top align="right">
								<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								  </td>
								</tr>
								<% 
  End Sub
  'ɾ�����
  Sub Delxc()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ�������!",ComeUrl):Response.End
	Conn.Execute("Delete From KS_Photoxc Where ID In(" & ID & ")")
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where xcid in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid=" & rs("id"))
	   KS.DeleteFile(rs("photourl"))
	   rs.movenext
	   loop
	end if
	Conn.execute("delete from ks_photozp where xcid in(" & id& ")")
	Conn.execute("delete from ks_uploadfiles where channelid=1028 and infoid in(" & id& ")")
	rs.close:set rs=nothing
	Call KSUser.AddLog(KSUser.UserName,"ɾ����������!",104)
	Response.Redirect ComeUrl
  End Sub
  'ɾ����Ƭ
  Sub Delzp()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ������Ƭ!",ComeUrl):Response.End
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where id in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   KS.DeleteFile(rs("photourl"))
	   Conn.execute("update ks_photoxc set xps=xps-1 where id=" & rs("xcid"))
	   rs.movenext
	   loop
	end if
	Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid in(" & id& ")")
	Conn.execute("delete from ks_photozp where id in(" & id& ")")
	Call KSUser.AddLog(KSUser.UserName,"ɾ������Ƭ����!",104)
	rs.close:set rs=nothing
	Response.Redirect ComeUrl
  End Sub
  '�ϴ���Ƭ
  Sub Addzp()
        Call KSUser.InnerLocation("�ϴ���Ƭ")
		  adddate=now:XCID=KS.ChkCLng(KS.S("XCID")):UserName=KSUser.GetUserInfo("RealName")
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				 if ($("input[name=pubtype][checked=true]").val()==1){
					if (document.myform.Title.value=="")
					  {
						alert("������������ƣ�");
						document.myform.Title.focus();
						return false;
					  }
				 }else if (document.myform.XCID.value==""){
					alert("��ѡ��������ᣡ");
					document.myform.XCID.focus();
					return false;
				  }	
				  	
				  var picSrcs='';
				  var src='';
				  $("#thumbnails").find(".pics").each(function(){
					 src=$(this).next().val().replace('|||','').replace('|','')+'@@@'+$(this).val()
					 if(picSrcs==''){
					  picSrcs=src;
					 }else{
					  picSrcs+='|'+src;
					 }
				  });
				  if (picSrcs==''){
				   alert('���ϴ���Ƭ!');
				   return false;
				  }
				  $('#PhotoUrls').val(picSrcs);
				 return true;  
				}
				</script>
	
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_Photo.asp?Action=AddSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>�� �� �� Ƭ</td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>ѡ����᣺</span></td>
                       <td width="88%">
					     <label><input type="radio" name="pubtype" value="0" onclick="$('#pub1').hide();$('#pub0').show();" checked>�����������</label>
					     <label><input type="radio" name="pubtype" value="1" onclick="$('#pub1').show();$('#pub0').hide();">���������</label>
						 <br/>
						 <div id="pub0" style="margin-top:5px">
						 <strong>ѡ�����</strong>
					   <select class="textbox" size='1' name='XCID' style="width:150">
							<% 
							Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							RS.Open "Select * From KS_Photoxc where username='" & KSUser.Username & "' order by id desc",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If XCID=RS("ID") Then
								  Response.Write "<option value=""" & RS("ID") & """ selected>" & RS("XCName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>	
						 
						 </div>
						 <div id="pub1" style="display:none;margin-top:5px">
						   <table border="0">
						    <tr>
							 <td>
						   <strong>������ƣ�</strong>
						     </td>
							 <td colspan="2">
						   <input class="textbox" name="Title" type="text" id="Title" style="width:300px; " value="<%=Title%>" maxlength="100" /><span style="color: #FF0000">*</span>
						     </td>
						   </tr>
						   <tr>
						     <td><strong>�����ܣ�</strong></td>
							 <td colspan="2"><textarea class="textbox" style="height:50px" name="Descript" cols="50" rows="5"></textarea></td>
						   </tr>
						   <tr>
						     <td><strong>�Ƿ񹫿���</strong></td>
							 <td><select onChange="if(this.options[selectedIndex].value=='3'){document.myform.all.mmtt.style.display='block';}else{document.myform.all.mmtt.style.display='none';}"  name="flag"><option value="1" selected>��ȫ����</option>
                      <option value="2">��Ա����</option>
                      <option value="3">���빲��</option>
                      <option value="4">��˽���</option>
                    </select></td><td><span class=child id=mmtt name="mmtt" style="display:none;">���룺<input type="password" name="password" style="width:120px" maxlength="16" value="" size="20"></span></td>
				           </tr>
						   <tr>
						     <td><strong>������</strong></td>
							 <td colspan="2"> <select class="textbox" size='1' name='ClassID' style="width:250"><% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_PhotoClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If ClassID=RS("ClassID") Then
								  Response.Write "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %></select></td>
							 </tr>
						  </table>
						   
						   
						 </div>
						 
						  <input class="textbox" name="PhotoUrls" type="hidden" id="PhotoUrls" style="width:350px; " maxlength="100" />
						   </td>
                    </tr>
                      

					<tr class="tdbg">
					  <td align="center"><span>�ϴ���Ƭ��</span></td>
					  <td style="padding-top:8px">
					  <style type="text/css">
			#thumbnails{background:url(../plus/swfupload/images/albviewbg.gif) no-repeat;min-height:200px;_height:expression(document.body.clientHeight > 200? "200px": "auto" );}
			#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:158px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
			#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px}

			</style>
			<link href="../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
			<script type="text/javascript" src="../ks_inc/kesion.box.js"></script>
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
				newImgDiv.innerHTML += '<div style="margin-top:10px;text-align:left">'+delstr+' <b>ע�ͣ�</b><input type="hidden" class="pics" id="pic'+pid+'" value="'+bigsrc+'"/><input type="text" name="picinfo'+pid+'" value="'+text+'" style="width:155px;" /></div>';
			
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
					upload_url: "swfupload.asp",
					post_params: {"BasicType":9997,"ChannelID":9997,"AutoRename":4},
	
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
					button_text : '<span class="button">�����ϴ�(��ͼ����2 MB)</span>',
					button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#666666;font-size: 14px; } ',
					button_text_top_padding: 3,
					button_text_left_padding: 0,
					button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
					button_cursor: SWFUpload.CURSOR.HAND,
					
					// Flash Settings
					flash_url : "../plus/swfupload/swfupload/swfupload.swf",
					flash9_url : "../plus/swfuploadswfupload/swfupload_FP9.swf",
	
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
	p.popup("<div style='text-align:left;padding-left:2px'>����ͼƬ��ַ</div>","<div style='padding:3px'>��http://��ͷ��Զ��ͼƬ��ַ,ÿ��һ��ͼƬ��ַ:<br/><textarea id='collecthttp' style='width:400px;height:150px'></textarea><br/><input type='button' value='ȷ ��' onclick='ProcessCollect()' class='button'/> <input type='button' value='ȡ ��' class='button' onclick='closeWindow()'/></div>",420);
	}
	function AddTJ(){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("<div style='text-align:left;padding-left:2px'>���ϴ��ļ���ѡ��</div>","<div style='padding:3px'><strong>ͼƬ��ַ:</strong><input type='text' name='x1' id='x1'> <input type='button' onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("ѡ��ͼƬ")%>&ChannelID=9997',550,290,window,$('#x1')[0]);\" value='ѡ��ͼƬ' class='button'/><br/><strong>��Ҫ����:</strong><input type='text' name='x3' id='x3'><br/><br/><input type='button' value='�� ��' onclick='ProcessAddTj()' class='button'/> <input type='button' value='ȡ ��' class='button' onclick='closeWindow()'/></div>",420);
	}
	function ProcessAddTj(){
	  if ($("#x1").val()==''){
	   alert('��ѡ��һ��ͼƬ��ַ!');
	   $("#x1").focus();
	   return false;
	  }
	  addImage($("#x1").val(),$("#x1").val(),$("#x3").val())
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
	   
	   var bigsrc=carr[0];
	   var smallsrc=carr[0];
	   addImage(bigsrc,smallsrc,'')
	 }
	 //$("#collecthttp").empty();
	 closeWindow();
	}
	</script>
	<table cellspacing="0" cellpadding="0">
		 <tr>
		  <td><div class="pn" style="margin: -6px 0px 0 0;"><span id="spanButtonPlaceholder"></span></div>
		 </td>
		 <td>
		<!-- <button type="button"  class="pn" onClick="OnlineCollect()" style="margin: -6px 0px 0 0;"><strong>���ϵ�ַ</strong></button>-->
		 <button type="button"  class="pn" onClick="AddTJ();" style="margin: -6px 0px 0 0;"><strong>ͼƬ��...</strong></button>
		 </td>
		 </tr>
		</table>

		<label><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)" checked="checked"/>��Ƭ���ˮӡ</label>
		<div id="divFileProgressContainer"></div>
		
		<div id="thumbnails"></div>

	   </td>
	   </tr>
	   
														 
                    <tr class="tdbg">
					  <td></td>
                      <td height="30">
					   <button id="button1" type="submit" class="pn"><strong>OK,��������</strong></button>
					</td>
                    </tr>
                  </form>
			    </table>
		  <%
  End Sub
    '�༭��Ƭ
  Sub Editzp()
        Call KSUser.InnerLocation("�༭��Ƭ")
		  Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_A_RS_Obj.Open "Select * From KS_PhotoZp Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not KS_A_RS_Obj.Eof Then
		     XCID  = KS_A_RS_Obj("XCID")
			 Title    = KS_A_RS_Obj("Title")
			 UserName   = KS_A_RS_Obj("UserName")
			 descript = ks_a_rs_obj("descript")
			 PhotoUrlS  = KS_A_RS_Obj("PhotoUrl")
		   End If
		   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.XCID.value=="0") 
				  {
					alert("��ѡ��������ᣡ");
					document.myform.XCID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("��������Ƭ���ƣ�");
					document.myform.Title.focus();
					return false;
				  }		
				 return true;  
				}
				
				</script>
				
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_Photo.asp?Action=EditSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>�� �� �� Ƭ</td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>ѡ����᣺</span></td>
                       <td width="88%"><select class="textbox" size='1' name='XCID' style="width:150">
                             <option value="0">-��ѡ�����-</option>
							  <% Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_Photoxc order by id desc",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If XCID=RS("ID") Then
								  Response.Write "<option value=""" & RS("ID") & """ selected>" & RS("XCName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>					  </td>
                    </tr>
                      <tr class="tdbg"  style="display:none">
                           <td  height="25" align="center"><span>��Ƭ���ƣ�</span></td>
                              <td><input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                        <span style="color: #FF0000">*
                                        <input class="textbox" name="PhotoUrls" type="hidden" id="PhotoUrls" style="width:350px; " maxlength="100" value="<%=photourls%>"/>
                                        </span></td>
                    </tr>
								<tr class="tdbg">
								  <td height="20" align="center">��ƬԤ����</td>
								  <td id="viewarea">
								    <table style='BORDER-COLLAPSE: collapse' borderColor='#c0c0c0' cellSpacing='1' cellPadding='2' border='1'><tr><td align='center' width='83' height='100' bgcolor='#ffffff'><img name='view1' width='83' height='100' src='<%=Photourls%>' title='��ƬԤ��'></td></tr></table> <input class="button" type='button' name='Submit3' value='ѡ����Ƭ��ַ...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("ѡ��ͼƬ")%>&ChannelID=9997',500,360,window,document.myform.PhotoUrls);" />
								</td>
				    </tr>
														 
								<tr class="tdbg">
                                   <td height="25" align="center"><span>��Ƭ���ܣ�</span></td>
                                  <td><textarea class="textbox" style="height:50px" name="Descript" cols="70" rows="5"><%=DESCRIPT%></textarea></td>
							  </tr>							 
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input type="submit" name="Submit"  class="button" value=" OK,�������� " />
                      <input type="reset" name="Submit2"   class="button" onClick="javascript:history.back()" value=" ȡ �� " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
  End Sub

   Sub EditSave()
    Dim RSObj,Descript,PhotoUrlArr,i
                 XCID=KS.ChkClng(KS.S("XCID"))
				 Title=Trim(KS.S("Title"))
				 UserName=Trim(KS.S("UserName"))
				 Descript=KS.S("Descript")
				 PhotoUrls=KS.S("PhotoUrls")
				 If PhotoUrls="" Then 
				    Response.Write "<script>alert('��û���ϴ���Ƭ!');history.back();</script>"
				    Exit Sub
				  End IF
				  on error resume next
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From KS_PhotoZP Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				  RSObj("Title")=left(Descript,200)
				  RSObj("XCID")=XCID
				  RSObj("PhotoUrl")=PhotoUrls
				  RSObj("Descript")=Descript
				  RSObj("PhotoSize") =KS.GetFieSize(Server.Mappath(replace(PhotoUrls,ks.getdomain,ks.setting(3))))
				RSObj.Update
				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(1029,KS.ChkClng(KS.S("ID")),PhotoUrls,1)
				 Call KSUser.AddLog(KSUser.UserName,"�޸�����Ƭ����! <a href=""" & PhotoUrls & """ target=""_blank"">�鿴</a>",104)
				 Response.Write "<script>alert('��Ƭ�޸ĳɹ�!');location.href='User_Photo.asp?Action=ViewZP&XCID=" & XCID& "';</script>"
  End Sub
  
  Sub AddSave()
    Dim RSObj,Descript,PhotoUrlArr,i,UpFiles,PhotoUrl,PubType,ClassID
	             PubType=KS.ChkClng(KS.S("PubType"))
                 XCID=KS.ChkClng(KS.S("XCID"))
				 ClassID=KS.ChkClng(KS.S("ClassID"))
				 Title=Trim(KS.S("Title"))
				 UserName=Trim(KS.S("UserName"))
				 Descript=KS.S("Descript")
				 PhotoUrls=KS.S("PhotoUrls")
				 If PhotoUrls="" Then 
				    Response.Write "<script>alert('��û���ϴ���Ƭ!');history.back();</script>"
				    Exit Sub
				  End IF
				 PhotoUrlArr=Split(PhotoUrls,"|")
				 
				  If XCID=0 And PubType=0 Then
				    Response.Write "<script>alert('��û��ѡ�����!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" And PubType=1 Then
				    Response.Write "<script>alert('��û�������������!');history.back();</script>"
				    Exit Sub
				  End IF
				
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				If PubType=1 Then
			        RSObj.Open "Select top 1 * From KS_Photoxc" ,conn,1,3
					  RSObj.AddNew
						RSObj("AddDate")=now
						if ks.SSetting(4)=1 then
						RSObj("Status")=0 '��Ϊ����
						else
						RSObj("Status")=1 '��Ϊ����
						end if
						RSObj("UserID")=KSUser.GetUserInfo("userid")
						RSObj("UserName")=KSUser.UserName
						RSObj("xcname")=Title
						RSObj("ClassID")=ClassID
						RSObj("Descript")=Descript
						RSObj("Flag")=KS.ChkClng(KS.S("Flag"))
						RSObj("Password")=KS.S("PassWord")
						RSObj("PhotoUrl")=Split(PhotoUrlArr(0),"@@@")(1)
					  RSObj.Update
					  RSObj.MoveLast
					  XCID=RSObj("id")
					   Call KS.FileAssociation(1028,XCID,RSObj("PhotoUrl"),0)
					   Call KSUser.AddLog(KSUser.UserName,"���������!����: "& Title & " <a href=""../space/?" & KSUser.GetUserInfo("userid") & "/showalbum/" & XCID & """ target=""_blank"">�鿴</a>",104)
				RSObj.Close
				End If
				
				RSObj.Open "Select top 1 * From KS_PhotoZP",Conn,1,3
				 For I=0 to ubound(PhotoUrlArr)
			    	RSObj.AddNew
					 PhotoUrl=Split(PhotoUrlArr(I),"@@@")(1)
					 RSObj("PhotoSize") =KS.GetFieSize(Server.Mappath(Replace(PhotoUrl,KS.GetDomain,KS.Setting(3))))
				     RSObj("Title")=left(Split(PhotoUrlArr(I),"@@@")(0),200)
				     RSObj("XCID")=XCID
					 RSObj("UserName")=KSUser.UserName
					 RSObj("PhotoUrl")=PhotoUrl
					 RSObj("Adddate")=Now
					 RSObj("Descript")=Split(PhotoUrlArr(I),"@@@")(0)
				   RSObj.Update
				   RSObj.MoveLast
				   Call KS.FileAssociation(1029,RSObj("ID"),PhotoUrlArr(i),0)
				 Next
				 RSObj.Close
				 Set RSObj=Nothing
				 
				 
				 Conn.Execute("update KS_Photoxc set xps=xps+" & Ubound(PhotoUrlArr)+1 & " where id=" & xcid)
				 Call KSUser.AddLog(KSUser.UserName,"�ϴ���" & Ubound(PhotoUrlArr)+1 & "����Ƭ�����! <a href=""../space/?" & KSUser.GetUserInfo("userid") & "/showalbum/" & xcid & """ target=""_blank"">�鿴</a>",104)
				 Response.Write "<script>if (confirm('��Ƭ����ɹ��������ϴ���?')){location.href='User_Photo.asp?Action=Add';}else{location.href='User_Photo.asp?Action=ViewZP&XCID=" & XCID& "';}</script>"
  End Sub

End Class
%> 
