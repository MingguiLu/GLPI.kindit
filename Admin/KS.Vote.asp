<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Vote
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Vote
        Private KS,KSCls
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMS20003") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Select Case KS.G("Action")
			 Case "Add","Edit" Call VoteAdd()
			 Case "Del" Call VoteDel()
			 Case "Set" Call VoteSet()
			 Case Else Call MainList()
			End Select
			
	  End Sub
	  
	  Sub MainList()
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<title>վ�����</title>"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "var Page='" & CurrentPage & "';" & vbCrLf
			.Write "</script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
			%>
			<script language="javascript">
			$(document).ready(function(){
				
		      $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
			  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		     })
			
			
			function VoteAdd()
			{
				location.href='KS.Vote.asp?Action=Add';
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=��������������� >> <font color=red>����µ�������</font>&ButtonSymbol=VoteAddSave';
			}
			function EditVote(id)
			{
			   if (id=='') id=get_Ids(document.myform);
			   if (id==''){
				 alert('��ѡ��Ҫ�༭�ĵ�������!');
				}else if(id.indexOf(',')==-1){
				location="KS.Vote.asp?Action=Edit&VoteID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=��������������� >> <font color=red>�༭��������</font>&ButtonSymbol=VoteEdit';
				}else{
				alert('һ��ֻ�ܱ༭һ����������!');
				}
			}
			function DelVote(id)
			{
			 if (id=='') id=get_Ids(document.myform);
			 if (id==''){
			   alert('����ѡ��Ҫɾ���ĵ�������!')
			 }else if  (confirm('���Ҫɾ��ѡ�еĵ���������?')){
			 location="KS.Vote.asp?Action=Del&Page="+Page+"&Voteid="+id;
			 }
			}
			function SetVoteNewest(id)
			{
				location="KS.Vote.asp?Action=Set&Page="+Page+"&Voteid="+id;
			}
			
			
			</script>
			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"">"
		    .Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""VoteAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��ӵ���</span></li>"
			.Write "<li class='parent' onclick=""EditVote('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭����</span></li>"
			.Write "<li class='parent' onclick=""DelVote('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ������</span></li>"
			.Write "</ul>"
			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""1"">"
			.Write "<form name='myform' action='KS.Vote.asp' method='post'>"
			.Write "<input type='hidden' name='action' value='Del'>"
			.Write "  <tr>"
			.Write "          <td width=""35"" height=""25"" class=""sort"">ѡ��</td>"
			.Write "          <td height=""25"" class=""sort""align=""center"">��������</td>"
			.Write "          <td width=""100"" class=""sort"" align=""center"">��̳����</td>"
			.Write "          <td width=""100"" class=""sort"" align=""center"">������</td>"
			.Write "          <td width=""120"" align=""center"" class=""sort"">ʱ��</td>"
			.Write "          <td width=""100"" class=""sort"" align=""center"">�Ƿ�����</td>"
			.Write "          <td width=""120"" class=""sort"" align=""center"">�������</td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_Vote order by NewestTF desc,AddDate desc"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					   .Write "<tr><td height='30' class='splittd' align='center' colspan='6'>��û����ӵ�������!</td></tr>"
					 Else
						        totalPut = RSObj.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								If CurrentPage > 1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								Call showContent
				End If
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			  Dim ID
			  With Response
					Do While Not RSObj.EOF
					   ID=RSObj("id")
					   .Write ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &ID & "' onclick=""chk_iddiv('" & ID & "')"">")
				       .Write ("<td class='splittd' align=center><input type='hidden' value='" & ID & "' name='VoteID'><input name='id'  onclick=""chk_iddiv('" & ID & "')"" type='checkbox' id='c"& ID & "' value='" & ID & "'></td>")
					  .Write "  <td class='splittd'  height='20'> &nbsp;&nbsp; <span VoteID='" & ID & "' ondblclick=""EditVote(this.VoteID)""><img src='Images/Vote.gif' align='absmiddle'>"
					  .Write    KS.GotTopic(RSObj("Title"), 50) & "</span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd'  align='center'>" 
					   If KS.ChkClng(RSOBj("TopicID"))=0 Then
					     .Write "��"
					   Else
					     .Write "<a href='" & KS.GetClubShowUrl(RSObj("TopicID")) & "' style='color:green' target='_blank'>��</a>"
					   End If
					  .Write " </td>"
					  .Write "  <td class='splittd'  align='center'>" & RSObj("UserName") & " </td>"
					  .Write "  <td class='splittd'  align='center'><FONT Color=red>" & RSObj("AddDate") & "</font> </td>"
					  If RSObj("NewestTF") = 1 Then
					   .Write "  <td class='splittd' align='center'><font color=red>��</font></td>"
					  Else
					   .Write "  <td class='splittd' align='center'>��</td>"
					  End If
					   .Write "  <td class='splittd' align='center'><a href=""javascript:EditVote('"&Id&"');"">�޸�</a> | <a href=""javascript:DelVote('"&Id&"');"" >ɾ��</a> | <a href=""../?do=vote&id=" & id & """ target=""_blank"">�鿴</a></td>"
					  .Write "</tr>"

					I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  Conn.Close
					 .Write "</table><table width='100%'><tr><td><div style='margin:5px'><b>ѡ��</b><a href='javascript:void(0)' onclick='Select(0)'>ȫѡ</a> -  <a href='javascript:void(0)' onclick='Select(1)'>��ѡ</a> - <a href='javascript:void(0)' onclick='Select(2)'>��ѡ</a> <input type='submit' class='button' value='ɾ ��' onclick=""return(confirm('ȷ��ɾ��ѡ�еĵ���������?'))""></form></td><td height='26' colspan='2' align='right'>"
					 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			  End With
			End Sub
			
			Sub VoteDel()
			 Dim ID,IDArr,I
			 ID=KS.S("VoteID")
			 If KS.IsNul(ID) Then Call KS.AlertHintScript("��ѡ��Ҫɾ��������!")
			 IDArr=Split(KS.FilterIds(ID),",")
			 For I=0 To Ubound(IDArr)
			 KS.DeleteFile(KS.Setting(3)&"config/voteitem/vote_" & IDArr(i) &".xml")
			 Conn.Execute("delete from KS_Vote where ID="&Clng(IDArr(i)))
			 Conn.Execute("delete from KS_PhotoVote where channelid=-1 and InfoID='"&Clng(IDArr(i))&"'")
			 Next
			 Response.redirect "KS.Vote.asp?Page="&KS.G("Page")
			End Sub
			
			Sub VoteSet()
				conn.execute "Update KS_Vote set NewestTF=0 where NewestTF=1"
				conn.execute "Update KS_Vote set NewestTF=1 Where ID=" & Clng(KS.G("VoteID"))
				Response.Write "<script language='JavaScript' type='text/JavaScript'>alert('���óɹ���');location.href='KS.Vote.asp?Page=" & KS.G("Page") & "';</script>"

			End Sub
			
			Sub VoteAdd()
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.Write "<title>�������-�������</title>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "<script src=""../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
				.Write "</head>"
				.Write "<body topmargin=""0"" leftmargin=""0"">"
	
				.Write "  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
				.Write "        <tr>"
				.Write "          <td width=""44%"" height=""25"" class=""sort"">"
				.Write "          <div align=""center""><strong>�� �� �� �� �� ��</strong></div></td>"
				.Write "        </tr>"
				.Write "      </table>"
	           
			   dim timelimit,Title,VoteTime,NewestTF,rs,sql,voteid,ItemArr,VoteNumArr,i,XMLStr
			   Dim VoteType,timebegin,timeend,nmtp,AllowGroupID,ipnum,ipnumS,templateid,Status,editnum
			   timelimit=0:nmtp=0:Status=1:editnum=0:timebegin=now:timeend=dateadd("m",1,now)
			   templateid="{@TemplateDir}/ͶƱҳ.html"
			   
				
				Title=trim(request.form("Title"))
				VoteTime=trim(request.form("VoteTime"))
				if VoteTime="" then VoteTime=now()
				NewestTF=trim(request("NewestTF"))
				
				ItemArr=Split(request("item"),",")
				VoteNumArr=Split(Request("VoteNum"),",")
				
				if Title<>"" then
					sql="select top 1 * from KS_Vote Where ID=" & ks.chkclng(request("voteid"))
					Set rs= Server.CreateObject("ADODB.Recordset")
					rs.open sql,conn,1,3
					if rs.eof then
					rs.addnew
					 rs("TopicID")=0
					 rs("VoteNums")=0
					end if
					rs("Title")=Title
					rs("timelimit")=KS.ChkClng(KS.G("TimeLimit"))
					If IsDate(Request("TimeBegin")) Then
					rs("TimeBegin")=Request("TimeBegin")
					Else
					rs("TimeBegin")=Now
					End If
					If IsDate(Request("TimeEnd")) Then
					 rs("TimeEnd")=Request("TimeEnd")
					Else
					 rs("TimeEnd")=Now
					End If
					rs("nmtp")=KS.ChkClng(Request("nmtp"))
					rs("groupids")=request.form("allowgroupid")
					rs("ipnum")=KS.ChkClng(Request.Form("ipnum"))
					rs("ipnums")=KS.ChkClng(Request.Form("ipnums"))
					rs("templateid")=request.form("templateid")
					rs("status")=KS.ChkClng(Request.Form("status"))
					rs("AddDate")=VoteTime
					rs("VoteType")=request("VoteType")
					rs("UserName")=KS.C("AdminName")
					if NewestTF="" then NewestTF=0
					rs("NewestTF")=NewestTF
					rs.update
					rs.movelast
					voteid=rs("id")
					rs.close
					if NewestTF=1 then conn.execute "Update KS_Vote set NewestTF=0 where NewestTF=1 and id<>" & voteid

					
					XMLStr="<?xml version=""1.0"" encoding=""gb2312"" ?>" &vbcrlf
					XMLStr=XMLStr&" <vote>" &vbcrlf
					for i=0 to ubound(ItemArr)
					  if trim(Itemarr(i))<>"" Then
					    XMLStr=XMLStr & "  <voteitem id=""" & i+1 &""">"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[" & Itemarr(i) &"]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <num>" & KS.ChkClng(VoteNumArr(i)) &"</num>" &vbcrlf
					    XMLStr=XMLStr & "  </voteitem>"&vbcrlf
						
					  End If
					Next
					XMLStr=XMLStr &" </vote>" &vbcrlf
					Call KS.WriteTOFile(KS.Setting(3) & "config/voteitem/vote_" & voteid & ".xml",xmlstr)
			        Application(KS.SiteSN&"_Configvote_"&voteid)=null
					set rs=nothing
					call CloseConn()
					if ks.chkclng(request("voteid"))=0 then
					 ks.die "<script>if (confirm('��ϲ��ͶƱ��Ŀ��ӳɹ������������')){location.href='KS.Vote.asp?action=Add';}else{location.href='KS.Vote.asp';}</script>"
					else
					 ks.die "<script>alert('��ϲ��ͶƱ��Ŀ�޸ĳɹ�!');location.href='KS.Vote.asp';</script>"
					end if
				end if
				 End With
				 
			if KS.ChkClng(request("voteid"))<>0 Then
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select top 1 * from ks_vote where id=" & KS.ChkClng(request("voteid")),conn,1,1
			  If Not RS.Eof Then
			    title    = RS("Title")
				VoteType = RS("VoteType")
				NewestTF = RS("NewestTF")
				timelimit= RS("timelimit")
				timebegin= RS("timebegin")
				timeEnd  = RS("timeEnd")
				nmtp     = RS("nmtp")
				AllowGroupID = RS("GroupIDs")
				ipnum    = RS("ipnum")
				ipnumS   = RS("ipnumS")
				templateid = RS("templateid")
				status=rs("status")
			  End If
			End If	 
				 
				%>
	

				<form method="POST" name="voteform" action="KS.Vote.asp?Action=Add">
				<input type="hidden" name="voteid" value="<%=request("voteid")%>">
						<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="ctable">
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>�������ƣ�</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<input name="Title" type="text" size="40" value="<%=title%>" maxlength="50">
							�磺��Ա�վ����Щ��Ŀ�ϸ���Ȥ!</td>
						  </tr>
                          <tr class="tdbg"> 
							<td height="25" align="right" class="clefttitle"><strong>�������ͣ�</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
										<select name="VoteType" id="VoteType">
											<option value="Single"<%If VoteType="Single" Then Response.Write " selected"%>>��ѡ</option>
											<option value="Multi"<%If VoteType="Multi" Then Response.Write " selected"%>>��ѡ</option>
									</select>
										<input name="NewestTF" type="checkbox" id="NewestTF" value="1"<%If NewestTF="1" Then Response.Write " checked"%> />
	��Ϊ���µ���</td>
						  </tr>						  
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>ͶƱ��Ŀ��</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							
							 <table border="0" cellpadding="0" cellspacing="0" style="margin-left:5px;" width="80%">
     
                 <tr>
                  <td colspan="3" height="30px">
							ͶƱ��չ����: 
						  <input name="vote_num" type="text" id="votenum" value="1" size="5" style="text-align:center"> 
						  <input type="button" name="Submit52" value="����ѡ��" class="button" onclick="javascript:doadd(jQuery('#votenum').val());"> 
							  
							  </td>
							 </tr>
							 <tr bgcolor='#DBEAF5'>
							 <td width='9%' height='20'> <div align='center'>���</div></td>
							 <td width='65%'> <div align='center'>��Ŀ����</div></td>
							 <td style='width: 100px'> <div align='center'>ͶƱ��</div></td>
							 </tr>
							 <tr>
							  <td colspan="3" id="addvote">
							  <%if request("voteid")<>"" then
							    Dim VoteXML,TaskNode,Node,N,TaskUrl,Taskid,Action
								set VoteXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
								VoteXML.async = false
								VoteXML.setProperty "ServerHTTPRequest", true 
								VoteXML.load(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" & request("voteid")& ".xml"))
								Dim TempStr
								editnum=VoteXml.DocumentElement.SelectNodes("voteitem").length
								 For Each Node In VoteXml.DocumentElement.SelectNodes("voteitem")
								  tempstr=tempstr & "<tr><td width=9% height=20> <div align=center><input type=hidden name=id value=" & Node.getAttribute("id") & ">" & Node.getAttribute("id") & "</div></td><td width='65%'> <div align=center><input type=text name=item size=40 value='" & trim(Node.childNodes(0).text) & "'></div></td><td width='26%'> <div align=center><input type=text name=votenum style=text-align:center value='" & Node.childNodes(1).text & "' size=6></div></td></tr>"
								 Next
							    end if
								response.write "<table width=100% border=0 cellspacing=1 cellpadding=3>"
								response.write tempstr
								response.write "</table>"
							  %>
							  
							  
							  
							  </td>
							 </tr>
							</table>
							<input name="editnum" type="hidden" id="editnum" value="<%=editnum%>"> 

							<script type="text/javascript">
    function doadd(num)
    {var i;
    var str="";
    var oldi=0;
    var j=0;
    oldi=parseInt(jQuery('#editnum').val());
    for(i=1;i<=num;i++)
    {
    j=i+oldi;
    str=str+"<tr><td width=9% height=20> <div align=center><input type=hidden name=id value=0>"+j+"</div></td><td width=65%> <div align=center><input type=text name=item size=40></div></td><td width=26%> <div align=center><input type=text name=votenum style='text-align:center' value=0 size=6></div></td></tr>";
    }
    window.addvote.innerHTML+="<table width=100% border=0 cellspacing=1 cellpadding=3>"+str+"</table>";
        jQuery('#editnum').val(j);
    }
	<%If request("voteid")="" Then%>
	doadd(8);
	<%end if%>
    </script>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>����ʱ�����ƣ�</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<label><input type='radio' name='timelimit' onclick="$('#time').hide();" value='0'<%IF timelimit="0" Then Response.Write " checked"%>>������</albe>
							<label><input type='radio' name='timelimit' onclick="$('#time').show();" value='1'<%IF timelimit="1" Then Response.Write " checked"%>>����</label>
							</td>
						  </tr>
						  <tbody id='time'<%if timelimit="0" then response.write " style='display:none'"%>>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>ʱ�����ƣ�</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							��Ч�� ��<input type='text' name='timebegin' value='<%=timebegin%>'>��
							<input type='text' name='timeend' value='<%=timeend%>'> ʱ���ʽΪ:YYYY-MM-DD HH:mm:ss
							</td>
						  </tr>
						  </tbody>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>����ͶƱ��</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<label><input type='radio' name='nmtp' value='0'<%If nmtp="0" Then Response.Write " checked"%>>��������ͶƱ</label>
							<label><input type='radio' name='nmtp' value='1'<%If nmtp="1" Then Response.Write " checked"%>>ֻ�����ԱͶƱ</label>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>�޶��û��飺</strong>
							<br/>�������벻Ҫѡ
							</td>
							<td colspan="3" bgcolor="#EEF8FE">
							<%=KS.GetUserGroup_CheckBox("AllowGroupID",AllowGroupID,5)%>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>ͬһIP��</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							һ����������Ͷ<input type="text" name='ipnum' value='<%=IPNUM%>' size='3' style='text-align:center'>�� ,�ܹ�����Ͷ<input type="text" name='ipnums' value='<%=IPNums%>' size='3' style='text-align:center'>�Ρ�tips:������������0
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>ͶƱҳģ�壺</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							 <input type="text" name="templateid" value="<%=templateid%>" size="40" id="templateid">
							 <%=KSCls.Get_KS_T_C("document.getElementById('TemplateID')")	%>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>״̬��</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<label><input type='radio' name='status' value='0'<%if status="0" then response.write " checked"%>>�ر�</label>
							<label><input type='radio' name='status' value='1'<%if status="1" then response.write " checked"%>>����</label>
							</td>
						  </tr>
									
									
							  </table>
							</form>
						</td>
					</tr>
	</table>
	<br/>
	<script>
	 function CheckForm()
	 { var form=document.voteform;
	  if (form.Title.value=='')
	   {
		 alert('�������������!');
		  form.Title.focus();
		 return false;
	   }
	   $("input[name='item']").each(function(){
	     $(this).val($(this).val().replace(/,/g,'��'));
	   });
	  document.voteform.submit();
	 }
	</script>
<%
			End Sub
			
End Class
%>
 
