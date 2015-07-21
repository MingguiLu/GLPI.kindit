<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New User_myask
KSCls.Kesion()
Set KSCls = Nothing

Class User_myask
        Private KS,KSUser
		Private CurrentPage,totalPut,i,PageNum
		Private RS,MaxPerPage,SQL,tablebody,action
		Private ComeUrl,TotalPages
		Private Sub Class_Initialize()
			MaxPerPage =10
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
		  KS.Die "<script>top.location.href='Login';</script>"
		End If
		Action=Request("action")
		
		If KS.S("Action")="cancel" Then	 Call FavCancel() : KS.Die ""
		
		CurrentPage=KS.ChkClng(Request("page"))
		if CurrentPage=0 Then CurrentPage=1
		Call KSUser.Head()
		Call KSUser.InnerLocation("我发表的主题")
		KSUser.CheckPowerAndDie("s19")
		call info()

	  End Sub

		
	  sub info()
		%>
	
			
		<div class="tabs">	
			<ul>
				<li<%If action="" then KS.Echo " class='select'"%>><a href="?">我的主题</a></li>
				<li<%If action="cy" Then KS.Echo " class='select'"%>><a href="?action=cy">参与的主题</a></li>
				<li<%If action="fav" Then KS.Echo " class='select'"%>><a href="?action=fav">我的收藏</a></li>
			</ul>
		</div>
			<table height='400' width="99%" align="center">
			<tr>
			<td valign="top">
		
   <%
          select Case Action
		   case "fav" fav
		   case else quesion
		  end select
		  
    Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
   %>
			 </td>
			 </tr>
		    </table>
		<%
			if request("action")="cy" then
	  ks.echo "<div style='color:red'><strong>说明：</strong>我参与的主题最多列出当前数据表的200条记录。</div>"
	end if

	end sub
	
	Sub Quesion()
	%>
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
			<tr height="28" class="titlename">
				<td height="25" align="center">主题</td>
				<td height="25" align="center">版块</td>
				<td width="10%" align="center">回复</td>
				<td width="15%" align="center">最后发表</td>
			</tr>
		<% 
		   dim 	sql

		
			'取帖子存放数据表
			if request("action")="cy" then
				Dim Nodes,Doc,TableName
				set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				Doc.async = false
				Doc.setProperty "ServerHTTPRequest", true 
				Doc.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
				Set Nodes=Doc.DocumentElement.SelectSingleNode("item[@isdefault='1']")
				TableName=nodes.selectsinglenode("tablename").text
				Set Doc=Nothing
				sql="select * from KS_Guestbook where id in(select top 200 topicid from " & TableName & " where Username='"&KSUser.UserName&"') order by LastReplayTime desc"
			else
			    sql="select * from KS_Guestbook where Username='"&KSUser.UserName&"' order by addTime desc"
			end if

		
			set rs=server.createobject("adodb.recordset")
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=4 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有发表过任何主题！</td>
			</tr>
		<%else
		          totalPut = RS.RecordCount
				  If CurrentPage < 1 Then	CurrentPage = 1
								
			      If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrentPage - 1) * MaxPerPage
				  Else
					  CurrentPage = 1
				  End If
				  i=0
		      do while not rs.eof
		%>
						<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
							<td height="25" class="splittd">
							<div class="ContentTitle">
							<a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank"><%=rs("subject")%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>发表时间:[<%=KS.GetTimeFormat1(rs("addtime"),false)%>]
							  状态:[<%if rs("verific")="1" then response.write "已审核" else response.write "未审核"%>]
							 </span>
							 </div>
							</td>
                            <td class="splittd" align="center">
							<%
							Dim Node
							KS.LoadClubBoard
			               Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & rs("boardid") &"]")
						   if not node is nothing then
						     KS.Echo "<a href='" & KS.GetClubListUrl(rs("boardid")) &"' target='_blank'>" & Node.SelectSingleNode("@boardname").text & "</a>"
						   else
						     KS.Echo "---"
						   end if
						   Set Node=Nothing
							%>
							</td>
							<td class="splittd" align=center>
							<%=RS("TotalReplay")%>
							</td>
							<td class="splittd" align=center>
							<a href='../space/?<%=RS("LastReplayUser")%>' target='_blank'><%=RS("LastReplayUser")%></a>
							<div class="Contenttips"><%=KS.GetTimeFormat1(RS("LastReplayTime"),True)%></div>
							</td>
						</tr>
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			loop
			end if
			rs.close
			set rs=Nothing
		%>
</table>
	<%
	End Sub
	
	
	Sub Fav()
	%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr height="28" class="title">
				<td height="25" align="center">主题</td>
				<td height="25" align="center">版块</td>
				<td width="10%" align="center">回复</td>
				<td width="15%" align="center">最后发表</td>
			</tr>
			<form name="myform" action="?action=cancel" method="post">
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select a.*,f.favorid from KS_Guestbook a inner join KS_AskFavorite f on a.id=f.topicid where f.Username='"&KSUser.UserName&"' order by LastReplayTime desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有收藏问题！</td>
			</tr>
		<%else
		
		                       totalPut = RS.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									i=0
		      do while not rs.eof
		%>
						
					<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
							<td height="25" class="splittd">
							<div class="ContentTitle">
							<input type="checkbox" name="favorid" value="<%=rs("favorid")%>">
							·<a href="<%=KS.GetClubShowUrl(rs("id"))%>" target="_blank"><%=rs("subject")%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>发表时间:[<%=KS.GetTimeFormat1(rs("addtime"),false)%>]
							  状态:[<%if rs("verific")="1" then response.write "已审核" else response.write "未审核"%>]
							 </span>
							 </div>
							</td>
                            <td class="splittd" align="center">
							<%
							Dim Node
							KS.LoadClubBoard
			               Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & rs("boardid") &"]")
						   if not node is nothing then
						     KS.Echo "<a href='" & KS.GetClubListUrl(rs("boardid")) &"' target='_blank'>" & Node.SelectSingleNode("@boardname").text & "</a>"
						   else
						     KS.Echo "---"
						   end if
						   Set Node=Nothing
							%>
							</td>
							<td class="splittd" align=center>
							<%=RS("TotalReplay")%>
							</td>
							<td class="splittd" align=center>
							<a href='../space/?<%=RS("LastReplayUser")%>' target='_blank'><%=RS("LastReplayUser")%></a>
							<div class="Contenttips"><%=KS.GetTimeFormat1(RS("LastReplayTime"),True)%></div>
							</td>
						</tr>	
						
						
						
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			
			loop
			end if
			rs.close
			set rs=Nothing
		%>
		<tr>
		 <td><input type="submit" value="取消收藏" class="button" onClick="return(confirm('确定取消收藏吗?'))"></td>
		</tr>
		</form>
	 </table>
	 <%
	End Sub
		
	Sub FavCancel()
		 Dim FavorID:Favorid=KS.FilterIDS(KS.S("favorid"))
		 if FavorID="" Then KS.AlertHintScript "对不起,您没有选择记录!"
		 Conn.Execute("Delete From KS_AskFavorite Where Favorid in(" & Favorid & ") and username='" & KSUser.UserName & "'")
		 Response.Redirect ComeUrl
	End Sub	
End Class
%> 
