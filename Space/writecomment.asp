<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS,KSUser
Set KS=New PublicCls
Set KSUser = New UserCls
Call KSUser.UserLoginChecked()
Dim ChannelID,RS,CommentStr,Total,UserIP

select case KS.S("Action")
  case "CommentSave"
    call CommentSave()
  case else
    Response.Write("document.write('" & GetWriteComment(KS.ChkClng(KS.S("UserID")),KS.S("ID"),KS.S("Title"),KS.S("UserName")) & "');")
end select


		'*********************************************************************************************************
		'��������GetWriteComment
		'��  �ã�ȡ�÷���������Ϣ
		'��  ����ID -��ϢID
		'*********************************************************************************************************
		Function GetWriteComment(UserID,ID,Title,UserName)
		%>
		function insertface(Val)
	      {  
		  if (Val!=''){ document.getElementById('Content').focus();
		  var str = document.selection.createRange();
		  str.text = Val; }
          }
		  function success()
			{
				var loading_msg='\n\n\t���Եȣ������ύ����...';
				var content=document.getElementById('Content');
				
				if (loader.readyState==1)
					{
						content.value=loading_msg;
					}
				if (loader.readyState==4)
					{   var s=loader.responseText;
						if (s=='ok')
						 {
						 alert('��ϲ,��������ѳɹ��ύ��');
						  location.reload();
						 }
						else
						 {alert(s);
						 }
					}
			}
		

		   function checkform()
		   { 
		    if (document.getElementById('AnounName').value=='')
			{
			 alert('�������ǳ�!');
			 document.getElementById('AnounName').focus();
			 return false;
			}
		    if (document.getElementById('Content').value=='')
			{
			 alert('��������������!');
			 document.getElementById('Content').focus();
			 return false;
			}
		   ksblog.ajaxFormSubmit(document.form1,'success')
           }
		   
		function ShowLogin()
		{ 
		 new KesionPopup().popupIframe('��Ա��¼','<%=KS.Setting(3)%>user/userlogin.asp?Action=Poplogin',397,184,'no');
		}
		<%
		If KS.SSetting(25)="0" And KS.IsNul(KS.C("UserName")) Then
		  GetWriteComment="<div style=""margin:20px""><strong>��ܰ��ʾ��</strong>ֻ�л�Ա�ſ��Է�������,����ǻ�Ա����<a href=""javascript:ShowLogin()"">��¼</a>,���ǻ�Ա����<a href=""../?do=reg"" target=""_blank"">ע��</a>��</div>"
		Else
		 GetWriteComment = "<table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 GetWriteComment = GetWriteComment & "<form name=""form1"" action=""WriteComment.asp?action=CommentSave"" method=""post"">"
		 GetWriteComment = GetWriteComment & "<input type=""hidden"" value=""" & UserID & """ name=""UserID""><input type=""hidden"" value=""" & UserName & """ name=""UserName""><input type=""hidden"" value=""" & ID & """ name=""ID"">"
		 GetWriteComment = GetWriteComment & "<tr><td colspan=""2"" height=""30"" class=""comment_write_title""><strong>��������:</strong>"
		 Dim HomePage
		 If KS.C("UserName")<>"" Then
		  HomePage=KS.Setting(2) & "/space/?" & KS.C("UserID")
		 Else
		  HomePage="http://"
		 End If
		GetWriteComment = GetWriteComment & "<br/>�ǳƣ�"
		GetWriteComment = GetWriteComment & "   <input name=""AnounName"" maxlength=""100"" type=""text"" id=""AnounName"" value=""" & KS.C("username") & """"
		If KS.C("UserName")<>"" Then GetWriteComment = GetWriteComment & " readonly"
		GetWriteComment = GetWriteComment & " style=""color:#999;width:35%;border:1px solid #ccc;background:#FBFBFB;""/><br/>��ҳ��"
		GetWriteComment = GetWriteComment & "    <input name=""HomePage"" maxlength=""150"" value=""" & HomePage & """ type=""text"" id=""HomePage"" style=""color:#999;width:55%;border:1px solid #ccc;background:#FBFBFB;"" /><br/>���⣺"
		GetWriteComment = GetWriteComment & "    <input name=""Title"" maxlength=""150"" value=""Re:" & Title & """ type=""text"" id=""Title"" style=""color:#999;width:55%;border:1px solid #ccc;background:#FBFBFB;"" /><input type=""hidden"" value=""" & Title & """ name=""OriTitle""></td>"
		GetWriteComment = GetWriteComment & "  </tr>"
		
		
		GetWriteComment = GetWriteComment & "  <tr>"
		GetWriteComment = GetWriteComment & "    <td height=""25"" width=""70%"" align=""center""><textarea name=""Content"" rows=""6"" id=""Content"" cols=""70"" style=""color:#999;width:98%;border:1px solid #ccc;background:#FBFBFB;overflow:auto""></textarea></td>"
		
		 Dim str:str="����|Ʋ��|ɫ|����|����|����|����|����|˯|���|����|��ŭ|��Ƥ|����|΢Ц|�ѹ�|��|�ǵ�|ץ��|��|"
		 Dim strArr:strArr=Split(str,"|")
		  GetWriteComment = GetWriteComment & "<td width=""140"">"
		 For K=0 to 19
		   GetWriteComment = GetWriteComment & "<img style=""cursor:pointer"" title=""" & strarr(k) & """ onclick=""insertface(\'[e" & k &"]\')""  src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">&nbsp;"
		   If (K+1) mod 5=0 Then GetWriteComment = GetWriteComment & "<br />"
		 Next

		GetWriteComment = GetWriteComment & "</td>"
		GetWriteComment = GetWriteComment & "  </tr>"
		GetWriteComment = GetWriteComment & "  <tr>"
		
		GetWriteComment = GetWriteComment & "    <td colspan=""2"" style=""text-align:left""><input type=""button"" onclick=""return(checkform())"" name=""SubmitComment"" id=""SubmitComment""class=""btn"" value=""�ύ����""/>"
		
		GetWriteComment = GetWriteComment & "    </td>"
		GetWriteComment = GetWriteComment & "  </tr>"
		GetWriteComment = GetWriteComment & "  </form>"
		GetWriteComment = GetWriteComment & "</table>"
		End If
		End Function  
  
        Sub CommentSave()
	    	Dim ID,UserName,HomePage,Content,Anonymous,Title
			ID=KS.ChkClng(KS.S("ID"))
			AnounName=KS.S("AnounName")
			HomePage=KS.S("HomePage")
			Content=KS.S("Content")
			Title=KS.S("Title")
			If Title="" Then Title="�ظ���������"
			IF ID="0" Then 
			 Response.Write("������������!")
			 Response.End
			End if
			if AnounName="" Then 
			 Response.Write("����д����ǳ�!'")
			 Response.End
			End if
			
			
			if Content="" Then 
			 Response.Write("����д��������!")
			 Response.End
			End if
			
			Set RS=Conn.Execute("Select top 1 UserName From KS_BlogInfo Where ID=" & ID)
			If RS.Eof And RS.Bof Then
			  RS.Close:Set RS=Nothing
			 Response.Write("������������!")
			 Response.End
			End If
			UserName=RS(0)
			RS.Close
			
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_BlogComment",Conn,1,3
			RS.AddNew
			 RS("LogID")=ID
			 RS("AnounName")=AnounName
			 RS("Title")=Title
			 RS("UserName")=UserName
			 RS("HomePage")=HomePage
			 RS("Content")=Content
			 RS("UserIP")=KS.GetIP
			 RS("AddDate")=Now
			RS.UpDate
			 RS.Close:Set RS=Nothing
			 Conn.Execute("Update KS_BlogInfo Set TotalPut=TotalPut+1 Where ID=" & ID)
			 
			 If KS.C("UserName")<>"" and  KS.S("From")<>"1" Then
			  Call KSUser.AddLog(KS.C("UserName"),"��<a href=""{$GetSiteUrl}space/?" & KS.S("UserID") &""" target=""_blank"">" & KS.S("UserName") & "</a>д�Ĳ���[<a href=""{$GetSiteUrl}space/?" & KS.S("UserID") & "/log/" & ID & """ target=""_blank"">" & KS.S("OriTitle") & "</a>]����������!",100)
			 End If
			 
			  Call CloseConn()
             If KS.S("From")="1" Then
			  Response.Write "<script>alert('������۷���ɹ�!');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
			 Else
			   response.write "ok"
			 End If
			 Set KS=Nothing
		End Sub
  
Call CloseConn
Set KS=Nothing
Set KSUser=Nothing
%>
