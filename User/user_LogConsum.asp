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
Set KSCls = New User_LogMoney
KSCls.Kesion()
Set KSCls = Nothing

Class User_LogMoney
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr
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
		
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		Call KSUser.InnerLocation("��ѯ�ҵ�ʹ�ü�¼")
		 If KS.S("page") <> "" Then
						          CurrentPage = CInt(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
							%>
		<div class="tabs">	
			<ul>
				<li><a href="user_logmoney.asp">�ʽ���ϸ</a></li>
				<li><a href="user_logpoint.asp">��ȯ��ϸ</a></li>
				<li><a href="user_logedays.asp">��Ч����ϸ</a></li>
				<li><a href="user_logscore.asp">������ϸ</a></li>
				<li class="select"><a href="user_LogConsum.asp">ʹ�ü�¼</a></li>
			</ul>
		</div>
			<div style="text-align:right;height:30px;line-height:30px;">  <a href='User_LogConsum.asp'><font color=red>�����м�¼</font></a> ��<a href='?d=1'>����[<%=conn.execute("select count(1) from KS_LogConsum where username='" & KSUser.UserName & "' and year(AddDate)=" & year(Now) & " and month(AddDate)=" & month(now) & " and day(AddDate)=" & day(now) &"")(0)%>]</a> ��<a href='?d=2'>2����[<%=conn.execute("select count(1) from KS_LogConsum where username='" & KSUser.UserName & "' and datediff(" &DataPart_D&",adddate," & SQLNowString &")<2")(0)%>]</a> ��<a href='?d=7'>һ����[<%=conn.execute("select count(1) from KS_LogConsum where username='" & KSUser.UserName & "' and datediff(" &DataPart_D&",adddate," & SQLNowString &")<7")(0)%>]</a>
		   </div>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
					<tr class=titlename align=middle>
					  <td width=80>�û���</td>
					  <td>����</td>
					  <td width=150 height="25">ʹ��ʱ��</td>
					  <td width=50>����</td>
					  <td>����</td>
					  <td>���</td>
					</tr>
					<%
					    Dim Param:Param=" Where l.channelid=i.channelid and UserName='" & KSUser.UserName &"'"
					    if ks.s("d")="1" then
						 Param=Param & " and year(l.AddDate)=" & year(Now) & " and month(l.AddDate)=" & month(now) & " and day(l.AddDate)=" & day(now)
						elseif KS.ChkClng(KS.S("d"))<>0 Then Param=Param & " and datediff(" &DataPart_D&",l.adddate," & SQLNowString &")<" & KS.ChkClng(KS.S("D"))
						end if
                        SqlStr="Select l.*,i.hits,i.tid,i.fname From KS_LogConsum l left join ks_iteminfo i on l.infoid=i.infoid " & Param & " order by logid desc"
						
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>�Ҳ�����Ҫ��ʹ�ü�¼!</td></tr>"
								 Else
									totalPut = RS.RecordCount
									If CurrentPage < 1 Then CurrentPage = 1
		
			
								If CurrentPage = 1 Then
									Call ShowContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowContent
									Else
										CurrentPage = 1
										Call ShowContent
									End If
								End If
				End If

						
						 %>
					
          </table>
		  </td>
		  </tr>
</table>
		  <%
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
  End Sub
    
  Sub ShowContent()
     on error resume next
     Dim I,intotalmoney,outtotalmoney
     Do While Not rs.eof 
	%>
    <tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
      <td  class="splittd" align=middle><%=rs("username")%></td>
      <td  class="splittd"><a href="<%=KS.GetItemURL(rs("ChannelID"),rs("Tid"),rs("InfoID"),rs("Fname"))%>" target="_blank"><%=rs("title")%></a></td>
      <td  class="splittd" align=middle><%=rs("adddate")%></td>
      <td   class="splittd" align=middle>
	  <% Select Case rs("basictype")
	      Case 1:Response.WRite "����"
		  Case 2:Response.Write "ͼƬ"
		  Case 3:Response.Write "����"
		  Case 4:Response.Write "����"
		  Case 7:Response.Write "ӰƬ"
		  Case 9:Response.Write "�Ծ�"
		 End Select
	 %>
	  </td>
      <td  class="splittd" align=center><%=KS.C_C(RS("Tid"),1)%></td>
      
      <td  class="splittd" align=center><%=rs("hits")%></td>
    </tr>
	<%
	   I = I + 1
	   RS.MoveNext
	  If I >= MaxPerPage Then Exit Do

	 loop
	%>
    
  </table>
		<%
		End Sub
  
End Class
%> 
