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
Set KSCls = New Admin_UserLog
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserLog
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,RS
		Private ID
		
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls= New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With KS
		 	    .echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
				.echo "<title>�û�ʹ�ü�¼����</title>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>" & vbCrLf
               Action=KS.G("Action")
				If Not KS.ReturnPowerResult(0, "KMUA10015") Then                 'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF

			 Page=KS.G("Page")
			 Select Case Action
			  Case "Del" ItemDelete
			  Case "DelAllRecord" DelAllRecord
			  Case Else MainList()
			 End Select
			.echo "</body>"
			.echo "</html>"
			End With
		End Sub
		
		Sub MainList()
			If Not IsEmpty(Request("page")) Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
		With KS
%>	   		
     <SCRIPT language=javascript>
		function DelDiggList()
		{
			var ids=get_Ids(document.myform);
			if (ids!='')
			 { 
				if (confirm('���Ҫɾ��ѡ�еļ�¼��?'))
				{
				$("#myform").action="KS.UserUseLog.asp?Action=Del&show=<%=KS.G("show")%>&ID="+ids;
				$("#myform").submit();
				}
			}
			else 
			{
			 alert('��ѡ��Ҫɾ��������!');
			}
		}
		function DelDigg()
		{
			if (confirm('���Ҫɾ��ѡ�еļ�¼��?'))
				{
				$("#myform").submit();
				}
		}
		function show(t,m,d)
		{
		new KesionPopup().PopupCenterIframe('�鿴����[<font color=red>'+t+'</font>]��¼','KS.UserUseLog.asp?action=list&infoid='+d,750,440,'auto')
		}
		function ShowCode(){
		new KesionPopup().PopupCenterIframe('�鿴Digg���ô���','KS.UserUseLog.asp?action=ShowCode',750,440,'no')
		}

		</SCRIPT>

	   <%
	
		.echo "</head>"
		
		.echo "<body topmargin='0' leftmargin='0'>"
		If KS.S("Action")="list" Then Call DiggDetail() : Exit Sub
		.echo "<div class='topdashed sort'>�鿴��Աʹ�ü�¼</div>"
		.echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.echo(" <form name=""myform"" id=""myform"" method=""Post"" action=""KS.UserUseLog.asp?Action=Del"">")
		.echo "    <tr class='sort'>"
		.echo "    <td width='30' align='center'>ѡ��</td>"
		.echo "    <td align='center'>����</td>"
		.echo "    <td align='center'>ʱ��</td>"
		.echo "    <td align='center'>�û���</td>"
		.echo "    <td align='center'>����</td>"
		.echo "    <td align='center'>����</td>"
		.echo "    <td align='center'>����</td>"
		.echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		   Dim Param:Param=" where L.ChannelID=I.ChannelID"
		   
				  SqlStr="Select l.*,i.hits,i.tid,i.fname From KS_LogConsum l inner join ks_iteminfo i on l.infoid=i.infoid " & Param & " order by logid desc"
				  RS.Open SqlStr, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				  .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=16 height='25' align='center'>û��ʹ�ü�¼!</td></tr>"
				 Else
					        totalPut = Conn.Execute("Select count(logid) from KS_LogConsum l inner join ks_iteminfo i on l.infoid=i.infoid " & Param)(0)
							If CurrentPage < 1 Then CurrentPage = 1
							
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent
			End If
		  .echo "  </td>"
		  .echo "</tr>"

		 .echo "</table>"
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo ("<tr><td width='170'><div style='margin:5px'><b>ѡ��</b><a href='javascript:Select(0)'><font color=#999999>ȫѡ</font></a> - <a href='javascript:Select(1)'><font color=#999999>��ѡ</font></a> - <a href='javascript:Select(2)'><font color=#999999>��ѡ</font></a> </div>")
		 .echo ("</td>")
	     .echo ("<td><input type=""button"" value=""ɾ��ѡ�еļ�¼"" onclick=""DelDiggList();"" class=""button""></td>")
	     .echo ("</form><td align='right'>")
	      Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.UserUseLog.asp", True, "��", CurrentPage,KS.QueryParam("page"))
	     .echo ("</td></tr></form></table>")
		 .echo ("<form action='KS.UserUseLog.asp?action=DelAllRecord' method='post' target='_hiddenframe'>")
		 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
		 .echo ("<div class='attention'><strong>�ر����ѣ� </strong><br>��վ������һ��ʱ���,��վ�Ļ�Ա ʹ�ü�¼����ܴ���Ŵ����ļ�¼,Ϊʹϵͳ���������ܸ���,����һ��ʱ�������һ�Ρ�")
		 .echo ("<br /> <strong>ɾ����Χ��</strong><input name=""deltype"" type=""radio"" value=1>10��ǰ <input name=""deltype"" type=""radio"" value=""2"" /> 1����ǰ <input name=""deltype"" type=""radio"" value=""3"" />2����ǰ <input name=""deltype"" type=""radio"" value=""4"" />3����ǰ <input name=""deltype"" type=""radio"" value=""5"" /> 6����ǰ <input name=""deltype"" type=""radio"" value=""6"" checked=""checked"" /> 1��ǰ  <input onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();"" type=""submit""  class=""button"" value=""ִ��ɾ��"">")
		 .echo ("</div>")
		 .echo ("</form>")
		End With
		End Sub
		Sub showContent()
		  Dim LogID
		  With KS
			 Do While Not RS.EOF
			 LogID=RS("LogID")
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & LogID & "' onclick=""chk_iddiv('" & LogID & "')"">"
		   .echo "<td class='splittd'><input name='id' onclick=""chk_iddiv('" & LogID & "')"" type='checkbox' id='c"& LogID & "' value='" & LogID & "'></td>"
		  .echo " <td class='splittd' height='22'><span style='cursor:default;'>"
		   .echo  RS("title")  & "</td>"
		   .echo " <td class='splittd' align='center'>" & RS("adddate") & " </td>"
		   .echo " <td class='splittd' align='center'>" & RS("username") & " </td>"
		   .echo " <td class='splittd' align='center'>" 
		   Select Case rs("basictype")
	      Case 1:Response.WRite "����"
		  Case 2:Response.Write "ͼƬ"
		  Case 3:Response.Write "����"
		  Case 4:Response.Write "����"
		  Case 7:Response.Write "ӰƬ"
		  Case 9:Response.Write "�Ծ�"
		 End Select
		   .echo "</td>"
		   .echo " <td class='splittd' align='center'>" & KS.C_C(RS("Tid"),1) & " </td>"
		   .echo " <td class='splittd' align='center'><a href='?action=Del&id=" & logid & "' onclick=""return(confirm('ȷ��ɾ����?'))"">ɾ��</a> </td>"
		   .echo "</tr>"
			I = I + 1:	If I >= MaxPerPage Then Exit Do
			RS.MoveNext
			Loop
		  RS.Close
		  End With
		 End Sub
		 
		 
		 Sub ItemDelete()
			Dim ID:ID = KS.G("ID")
			If ID="" Then KS.AlertHintScript "��û��ѡ��Ҫɾ���ļ�¼!"
			conn.Execute ("Delete From KS_LogConsum Where LogID IN(" & KS.FilterIds(ID) & ")")
		    response.redirect request.servervariables("http_referer") 
		 End Sub
		 
		
		 Sub DelAllRecord()
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>366"
		  End Select
   		  If Param<>"" Then Conn.Execute("Delete From KS_LogConsum Where " & Param)
          KS.echo "<script>$(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();alert('��ϲ,ɾ��ָ�������ڵļ�¼�ɹ�!');</script>"
		 End Sub
End Class
%> 
