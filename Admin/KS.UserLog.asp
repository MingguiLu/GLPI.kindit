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
        Private KS,Action,KSCls
		Private I, totalPut, Page,MaxPerPage, SqlStr,RS
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
				.echo "<title>�û���̬����</title>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>" & vbCrLf
               Action=KS.G("Action")
				If Not KS.ReturnPowerResult(0, "KMUA10014") Then                 'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF

			 Page=KS.ChkClng(KS.G("Page"))
			 If Page < 1 Then Page = 1
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
				$("#myform").action="KS.UserLog.asp?Action=Del&show=<%=KS.G("show")%>&ID="+ids;
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
		new KesionPopup().PopupCenterIframe('�鿴����[<font color=red>'+t+'</font>]��¼','KS.UserLog.asp?action=list&infoid='+d,750,440,'auto')
		}
		function ShowCode(){
		new KesionPopup().PopupCenterIframe('�鿴Digg���ô���','KS.UserLog.asp?action=ShowCode',750,440,'no')
		}

		</SCRIPT>

	   <%
	
		.echo "</head>"
		
		.echo "<body topmargin='0' leftmargin='0'>"
		If KS.S("Action")="list" Then Call DiggDetail() : Exit Sub
		.echo "<div class='topdashed sort'>��Ա��̬��¼</div>"
		.echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.echo(" <form name=""myform"" id=""myform"" method=""Post"" action=""KS.UserLog.asp?Action=Del"">")
		.echo "    <tr class='sort'>"
		.echo "    <td width='30' align='center'>ѡ��</td>"
		.echo "    <td align='center'>��̬</td>"
		.echo "    <td align='center'>ʱ��</td>"
		.echo "    <td width='8%' align='center'>���</td>"
		.echo "  </tr>"
		   Dim Param:Param=" 1=1"
		   If KS.G("Key")<>"" Then Param=" UserName='" & KS.S("Key") & "'"
		   
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
					Cmd("@tblName")="KS_UserLog"
					Cmd("@fldName")= "ID"
					Cmd("@pagesize")=MaxPerPage
					Cmd("@pageindex")=page
					Cmd("@ordertype")=1
					Cmd("@strWhere")=Param
					Cmd("@fieldIds")="*"
					Set Rs=Cmd.Execute
	 Else
			SQLStr=KS.GetPageSQL("KS_UserLog","ID",MaxPerPage,Page,1,Param,"*")
			Set RS = Server.CreateObject("AdoDb.RecordSet")
			RS.Open SQLStr, conn, 1, 1
	 End If
		   
		   
				  'SqlStr = "SELECT * From KS_UserLog " & Param & " order by ID Desc"
				 ' RS.Open SqlStr, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				  .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=6 height='25' align='center'>û�л�Ա��̬!</td></tr>"
				 Else
					totalPut = Conn.Execute("Select count(id) from KS_UserLog where " & Param)(0)
					Call showContent
			End If
		  .echo "  </td>"
		  .echo "</tr>"

		 .echo "</table>"
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo ("<tr><td width='170'><div style='margin:5px'><b>ѡ��</b><a href='javascript:Select(0)'><font color=#999999>ȫѡ</font></a> - <a href='javascript:Select(1)'><font color=#999999>��ѡ</font></a> - <a href='javascript:Select(2)'><font color=#999999>��ѡ</font></a> </div>")
		 .echo ("</td>")
	     .echo ("<td><input type=""button"" value=""ɾ��ѡ�еļ�¼"" onclick=""DelDiggList();"" class=""button""></td>")
	     .echo ("</form></td><td><form name='sform' action='?' method='post'><strong>���û���������</strong><input class='textbox' type='text' name='key'> <input class='button' type='submit' value='����'/></form></td></tr></table>")
	      Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
		 .echo ("<br /> <br /> <br /> <form action='KS.UserLog.asp?action=DelAllRecord' method='post' target='_hiddenframe'>")
		 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
		 .echo ("<div class='attention'><strong>�ر����ѣ� </strong><br>��վ������һ��ʱ���,��վ�Ļ�Ա��̬��¼����ܴ���Ŵ����ļ�¼,Ϊʹϵͳ���������ܸ���,����һ��ʱ�������һ�Ρ�")
		 .echo ("<br /> <strong>ɾ����Χ��</strong><input name=""deltype"" type=""radio"" value=1>10��ǰ <input name=""deltype"" type=""radio"" value=""2"" /> 1����ǰ <input name=""deltype"" type=""radio"" value=""3"" />2����ǰ <input name=""deltype"" type=""radio"" value=""4"" />3����ǰ <input name=""deltype"" type=""radio"" value=""5"" /> 6����ǰ <input name=""deltype"" type=""radio"" value=""6"" checked=""checked"" /> 1��ǰ  <input onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();"" type=""submit""  class=""button"" value=""ִ��ɾ��"">")
		 .echo ("</div>")
		 .echo ("</form>")
		End With
		End Sub
		Sub showContent()
		  With KS
			 Do While Not RS.EOF
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
		   .echo "<td class='splittd'><input name='id' onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		  .echo " <td class='splittd' height='22'><span style='cursor:default;'><img src='../images/user/log/" & rs("ico") & ".gif' align='absmiddle'>"
		   .echo  RS("username")  & Replace(RS("note"),"{$GetSiteUrl}",KS.GetDomain) & "</td>"
		   .echo " <td class='splittd' align='center'>" & RS("adddate") & " </td>"
		   .echo " <td class='splittd' align='center'><a href='?action=Del&id=" & rs("id") & "' onclick=""return(confirm('ȷ��ɾ����?'))"">ɾ��</a> </td>"
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
			conn.Execute ("Delete From KS_UserLog Where ID IN(" & KS.FilterIds(ID) & ")")
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
   		  If Param<>"" Then Conn.Execute("Delete From KS_UserLog Where " & Param)
          KS.echo "<script>$(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();alert('��ϲ,ɾ��ָ�������ڵļ�¼�ɹ�!');</script>"
		 End Sub
End Class
%> 
