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
Set KSCls = New UserReg
KSCls.Kesion()
Set KSCls = Nothing

Class UserReg
        Private KS, KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/UserFunction.asp"-->
		<%
       Public Sub loadMain()
		  Call KSUser.Head()
		  Call KSUser.InnerLocation("��Աע�ἤ��")
			Dim UserName:UserName=KS.S("UserName")
			Dim CheckNum:CheckNum=KS.S("CheckNum")
		  If KS.S("Flag")="Check" Then
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select UserName,RndPassWord,Email,CheckNum,locked,AllianceUser From KS_User Where UserName='" & UserName & "'",Conn,1,3
			If RS.Eof And RS.Bof Then
			rs.close:set rs=nothing
			 Response.Write "<script>alert('�Բ�����������û��������ڣ�');history.back();</script>":response.end
			else
			  if rs("checknum")<>checknum then
			  rs.close:set rs=nothing
			   Response.Write "<script>alert('�������������������룡');history.back();</script>":response.end
			  else
			   rs("locked")=0
			   rs.update
			   
			    Dim MailBodyStr,ReturnInfo
			    MailBodyStr = Replace(KS.Setting(147), "{$UserName}", rs("UserName"))
				MailBodyStr = Replace(MailBodyStr, "{$PassWord}", rs("RndPassWord"))
				MailBodyStr = Replace(MailBodyStr, "{$SiteName}", KS.Setting(0))
				ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-��Աע��ɹ�", RS("Email"),rs("UserName"), MailBodyStr,KS.Setting(11))

				IF ReturnInfo="OK" Then
				  ReturnInfo="<li>ע��ɹ�!�����û���:<font color=red>" & RS("UserName") & "</font>,�ѽ��û��������뷢����������!</li>"
				End If
				'���Ƽ��˼ӻ���
				Dim AllianceUser:AllianceUser=RS("AllianceUser")
				If AllianceUser<>RS("UserName") Then
				  If Not Conn.Execute("Select Top 1 UserID From KS_User Where UserName='" & AllianceUser & "'").eof Then
				   '�ж���û�ж����Ƽ�ע��,����ע��Ĳ�������
				   If Conn.Execute("Select top 1 * From KS_PromotedPlan Where UserIP='" & KS.GetIP & "' And DateDiff(" & DataPart_D & ",AddDate," & SqlNowString & ")<1 And UserName='" & AllianceUser & "'").eof Then
				   Call KS.ScoreInOrOut(AllianceUser,1,KS.ChkClng(KS.Setting(144)),"ϵͳ","�ɹ��Ƽ�һ��ע���û�:" & UserName & "!",0,0)
				   
				   Conn.Execute("Insert InTo KS_PromotedPlan(UserName,UserIP,AddDate,ComeUrl,Score,AllianceUser) values('" & AllianceUser & "','" & KS.GetIP & "'," & SqlNowString & ",'" & KS.URLDecode(Request.ServerVariables("HTTP_REFERER")) & "'," & KS.ChkClng(KS.Setting(144)) & ",'" & UserName & "')")
				   End If
				 End If
				End If
				rs.close:set rs=nothing
			   Response.Write "<script>alert('��ϲ��,�˺ż���ɹ�,�����ڿ���������¼�ˣ�');location.href='../user/login';</script>":response.end
			  end if
			end if
		  End If
		   %>
		    <form name="myform" method="post" action="?Flag=Check" onSubmit="return CheckForm();">
                 <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="border">
					  <tr class="title">
							<td height="24" align="center" colspan="2">�� �� �� ��</td>
					  </tr>
						  <TR class="tdbg">
						    <TD height=25 align="right">�����û�����</TD>
						    <TD><input name="UserName" type="text" id="UserName" size="20" value="<%=UserName%>"></TD>
			              </TR>
						  <TR class="tdbg">
							<TD width="40%" height=25 align="right"> ���ļ����룺</TD>
							<TD width="60%"><input name="CheckNum" type="text" id="CheckNum" size="20" value="<%=CheckNum%>"></TD>
						  </TR>
						  <TR class="tdbg">
							<TD  colspan="2" height=42 align="center"> 
							<input name="Submit" type="submit" class="button" value="ȷ������" style="padding:3px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </TD>
						  </TR>
						</TBODY>
					  </TABLE>
</form>
		   <%
		End Sub

End Class
%>

 
