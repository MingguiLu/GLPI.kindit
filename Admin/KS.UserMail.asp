<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_UserMail
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserMail
        Private KS
		Private Action,BodySQL,MailBody,TotalPut,CurrentPage,MaxPerPage
		Private Title, Content,sendername,senderemail, Numc,groupid,sendfile

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    If Not KS.ReturnPowerResult(0, "KMUA10009") Then
			  Response.Write "<script src='../ks_inc/jquery.js'></script>"
			  Response.Write ("<script>parent.frames['BottomFrame'].location.href='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
	        Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "<script src=""../ks_inc/Common.js"" language=""JavaScript""></script>"
			Response.Write "<script src=""../ks_inc/jquery.js"" language=""JavaScript""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"<table width=""100%""  height=""25"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			Response.Write " <tr>"
			Response.Write"	<td height=""25"" class=""sort"" style='text-align:left'> "
			Response.Write " &nbsp;&nbsp;<strong>������:</strong><a href='KS.UserMail.asp?action=MailSub'>�����ʼ�����</a> |  <a href='KS.UserMail.asp'>�����ʼ�</a> | <a href='?action=MailOut'>�����ʼ�</a>"
			Response.Write	" </td>"
			Response.Write " </tr>"
			Response.Write"</TABLE>"
		Action=Trim(Request("Action"))
		Select Case Action
		Case "Send"
			call send()
		Case "MailOut"
		    call MailOut()
		Case "DoExport"  '�������ı��ļ�
		    call DoExport()
	    Case "MailSub"
		    call MailSub()
		Case "active"
		    call active()
		Case "DelMail" 
		    call DelMail()
		Case else
			call sendmsg()
		end Select
		Response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 7.0, Copyright (c) 2006-2010 <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"%>
		</body>
		</html>
		<%
		End Sub
		
		Sub MailSub()
		 Dim RS,Param,SqlStr
		 CurrentPage=KS.ChkClng(KS.S("Page"))
		 If CurrentPage<=0 Then CurrentPage=1
		 MaxPerPage=20
		 With KS
		 .echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.echo(" <form name=""myform"" id=""myform"" method=""Post"" action=""KS.UserMail.asp"">")
		.echo " <input type='hidden' name='action' id='action' value='DelMail'/>"
		.echo " <input type='hidden' name='v' id='v' value='0'/>"
		.echo "    <tr class='sort'>"
		.echo "    <td width='30' align='center'>ѡ��</td>"
		.echo "    <td align='center'>�ʼ�</td>"
		.echo "    <td align='center'>�ύʱ��</td>"
		.echo "    <td align='center'>����Ȥ��Ŀ</td>"
		.echo "    <td width='8%' align='center'>״̬</td>"
		.echo "    <td align='center'>����</td>"
		.echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		          Param=" where 1=1"
				  SqlStr = "SELECT * From KS_UserMail " & Param & " order by ID Desc"
				  RS.Open SqlStr, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				  .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=6 height='25' align='center'>û�л�Ա�ύ����!</td></tr>"
				 Else
					        totalPut = Conn.Execute("Select count(id) from KS_UserMail" & Param)(0)
							If CurrentPage < 1 Then CurrentPage = 1
							
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent(RS)
			End If
		  .echo "  </td>"
		  .echo "</tr>"

		 .echo "</table>"
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo ("<tr><td width='170'><div style='margin:5px'><b>ѡ��</b><a href='javascript:Select(0)'><font color=#999999>ȫѡ</font></a> - <a href='javascript:Select(1)'><font color=#999999>��ѡ</font></a> - <a href='javascript:Select(2)'><font color=#999999>��ѡ</font></a> </div>")
		 .echo ("</td>")
	     .echo ("<td colspan=6><input type=""submit"" onclick=""$('#action').val('DelMail');return(confirm('�˲���������,ȷ��ɾ��ѡ�еļ�¼��'))"" value=""ɾ��ѡ�еļ�¼""  class=""button""> <input type=""submit"" onclick=""$('#action').val('active');$('#v').val(1)"" value=""��������""  class=""button""> <input type=""submit"" onclick=""$('#action').val('active');$('#v').val(0)"" value=""��������""  class=""button""></td>")
	     .echo ("</form></tr><tr><td  colspan=10 align='right'>")
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	     .echo ("</td></tr></form></table>")
		 End With
		End Sub
		
		Sub ShowContent(RS)
		  Dim i:i=0
		 With KS
			 Do While Not RS.EOF
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
		   .echo "<td class='splittd'><input name='id' onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		  .echo " <td class='splittd' height='22'><span style='cursor:default;'>"
		   .echo  RS("email")  &  "</td>"
		   .echo " <td class='splittd' align='center'>" & RS("adddate") & " </td>"
		   If KS.IsNul(rs("classid")) Then
		   .echo " <td class='splittd' align='center'> ȫ�� </td>"
		   Else
		       Dim ClassID,ClassIDArr,KK,classinfo
			   ClassID=Replace(rs("ClassID")," ","")
			   ClassIDArr=Split(ClassID,",")
			   classinfo=""
			   For kk=0 To Ubound(ClassIDArr)
				 If kk<>Ubound(ClassIDArr) Then
				  ClassInfo=ClassInfo & KS.C_C(ClassIDArr(kk),1) & "��"
				 Else
				  ClassInfo=ClassInfo & KS.C_C(ClassIDArr(kk),1) 
				 End If
			   Next
		   .echo " <td class='splittd' align='center' title='" & classinfo &"'> " & ks.gottopic(classinfo,30) & " </td>"
			   
		   End If
		   If rs("activetf")="0" then
		   .echo " <td class='splittd' align='center' style='color:red'> δ���� </td>"
		   else
		   .echo " <td class='splittd' align='center' style='color:green'> �Ѽ��� </td>"
		   end if
		   .echo " <td class='splittd' align='center'><a href='?action=active&v=1&id=" & rs("id") & "'>����</a> <a href='?action=active&v=0&id=" & rs("id") & "'>����</a> <a href='?action=DelMail&id=" & rs("id") & "' onclick=""return(confirm('ȷ��ɾ����?'))"">ɾ��</a> </td>"
		   .echo "</tr>"
			I = I + 1:	If I >= MaxPerPage Then Exit Do
			RS.MoveNext
			Loop
		  RS.Close
		  End With
		End Sub
		
		Sub DelMail()
		  Dim ID:ID=KS.G("ID")
		  id=ks.filterIds(id)
		  If KS.IsNul(ID) Then KS.AlertHintScript "û��ѡ��Ҫɾ��������!"
		  Conn.Execute("Delete From KS_UserMail Where ID In(" & ID & ")")
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		Sub active()
		  Dim ID:ID=KS.G("ID")
		  id=ks.filterIds(id)
		  If KS.IsNul(ID) Then KS.AlertHintScript "û��ѡ��Ҫɾ��������!"
		  Conn.Execute("update ks_usermail set activetf=" & ks.chkclng(ks.g("v")) & " Where ID In(" & ID & ")")
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		Sub SendMsg()
		%>
		<script src="../ks_inc/jquery.js"></script>
		<SCRIPT language=JavaScript>
function CheckForm(){
  if (document.myform.title.value==''){
     alert('�ʼ����ⲻ��Ϊ�գ�');
     document.myform.title.focus();
     return false;
  }
  if (parseInt($("input[name=sendtype][checked=true]").val())==1){
	   if (CKEDITOR.instances.Content.getData()=="" )
		{
		  alert("�ʼ����ݲ���Ϊ�գ�");
		  CKEDITOR.instances.Content.focus();
		  return false;
	   } 
 }
  return true;  
}
</SCRIPT>
</head>
<body><br>
 <% 
 dim InceptType:InceptType=KS.G("InceptType")
 if InceptType="" then InceptType="-1"
 dim userid:userid=KS.FilterIds(replace(request("userid")," ",""))
 dim usernamelist
 if userid<>"" then
	 dim rs:set rs=KS.InitialObject("adodb.recordset")
	 rs.open "select userid,username from ks_user where userid in("& userid & ")",conn,1,1
	 do while not rs.eof
	  if usernamelist="" then
	   usernamelist=rs(1)
	  else
	   usernamelist=usernamelist &"," & rs(1)
	  end if
	  rs.movenext
	 loop
	 rs.close:set rs=nothing
 end if
 %>
  <table class="ctable" cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
<FORM name=myform onSubmit="return CheckForm();" action=KS.UserMail.asp method=post>
    <tr class=sort>
      <td align=middle colSpan=2 height=22><B>�� �� �� ��</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">�ռ���ѡ��</td>
      <td>
        <table>
          <tr>
            <td colspan="2" style="color:red;font-weight:bold">
              <Input type=radio<%if InceptType="-1" then response.write " CHECKED"%> value="-1" name=InceptType> ���ж������䣬���� <font color=green><%=Conn.Execute("select count(1) from ks_usermail where activetf=1")(0)%></font> ����Ч����<span style='color:#999999'>��ͨ��<a href='../plus/mailsub.asp' target='_blank'>/plus/mailsub.asp</a>���ĵ��û���</span></td>
          </tr>
          <tr>
            <td>
              <Input type=radio<%if InceptType="0" then response.write " CHECKED"%> value="0" name=InceptType> ���л�Ա</td>
            <td></td>
          </tr>
          <tr>
            <td vAlign=top>
              <Input type=radio value="1" name=InceptType<%if InceptType="1" then response.write " CHECKED"%>> ָ����Ա��</td>
            <td><%=KS.GetUserGroup_CheckBox("GroupID",0,4)%></td>
          </tr>
          <tr>
            <td vAlign=top>
              <Input type=radio value="2" name=InceptType<%if InceptType="2" then response.write " CHECKED"%>> ָ���û���</td>
            <td>
              <Input size=40 name=inceptUser value="<%=usernamelist%>" class="textbox">
              ����û���������<font color=#0000ff>Ӣ�ĵĶ���</font>�ָ�</td>
          </tr>
          <tr>
            <td vAlign=top>
              <Input type=radio value="3" name=InceptType<%if InceptType="3" then response.write " CHECKED"%>>              ָ��Email</td>
            <td>
              <Input size=40 name=InceptEmail class="textbox"> 
              ���Email������<font color=#0000ff>Ӣ�ĵĶ���</font>�ָ�</td>
          </tr>
        </table>      </td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%" class="clefttitle">�ʼ����⣺</td>
      <td width="85%">
        <Input size="50" name=title id="title" class="textbox" value="�ʼ����ķ�������"> 
		<a href="javascript:void(0)" onClick="$('#title').val('')">���</a>
		</td>
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">�ʼ����ݣ�</td>
      <td>
	  <label style="color:#FF0000;font-weight:bold"><input type="radio" name="sendtype" checked="checked" onClick="$('#c0').show();$('#c1').hide();" value="0">�ʼ���������</label>
	  <label><input type="radio" name="sendtype" onClick="$('#c1').show();$('#c0').hide();" value="1">��ͨ����</label>
	  <script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
	  <div id="c0">
	    <table>
		 <tr>
		  <td colspan="2">
	    ����ƪ����<input type="text" name="sendnum" value="10" style="width:50px;text-align:center"/> ƪ  �����������ĵ���Ŀ��ѡ��ָ����ƪ�����͡�<br/>
	    �޶�������<input type="text" name="sendday" value="3" style="width:50px;text-align:center"/> ��  ������������0������ֻ��ָ��������ѡ�����·��͡�
		  </td>
		 </tr>
		 <tr>
		   <td valign="top" nowrap="nowrap"><strong>����ģ�壺</strong>
		  
		   
		   </td>
		   <td><textarea id="mailtemplate" name="mailtemplate" style="width:500px;height:200px"><strong>���ã���л�����롾{$GetSiteName}����վ�Ķ��ķ���</strong><br/>�����Ǳ��ڸ���������Ȥ����ĿΪ�����͵Ķ�������:<br/><br>
{$MailContent}<br>
<div style="text-align:right">{$GetSiteName}<br/>����ʱ�䣺{$SendDate}</div>
<a href="{$GetSiteUrl}" target="_blank">��˷��ʱ�վ</a> | <a onClick="return(confirm('��ȷ��ȡ���ö��ķ�����'))" href="{$GetCancelUrl}" target="_blank">ȡ�����ķ���</a><br>
		</textarea>
			    <script type="text/javascript">
                CKEDITOR.replace('mailtemplate', {width:"560",height:"200px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			   </script>
		
		 <br/>
		     ���ñ�ǩ��
			  {$GetSiteName}-��վ����,{$GetSiteUrl} -��վURL,{$MailContent} -��������,{$SendDate}-����ʱ��,{$GetCanCelUrl}-ȡ�����ĵ�URL
		
		</td>
	     </tr>
		</table>
	  </div>
	  <div id="c1" style="display:none">
	  <TEXTAREA id=Content style="display:none" name=Content></TEXTAREA> 
			    <script type="text/javascript">
                CKEDITOR.replace('Content', {width:"580",height:"200px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			   </script>
	  </div>
	  
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">ѡ�񸽼���</td>
      <td>����һ.<Input size=30 type="text" id="file1" name="sendfile"> <input class="button"  type='button' name='Submit' value='ѡ��...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=1&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,$('#file1')[0]);"><br />
	  ������.<Input size=30 type="text" id="file2" name="sendfile"> <input class="button"  type='button' name='Submit' value='ѡ��...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=1&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,$('#file2')[0]);"><br />
	  ������.<Input size=30 type="text" id="file3" name="sendfile"> <input class="button"  type='button' name='Submit' value='ѡ��...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=1&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,$('#file3')[0]);"><br />
	  
	  <font color=red>˵���������ļ��������������ƣ�������ܵ��·���ʧ�ܣ�</font>
	  </td>
    </tr>
	
    <tr class=tdbg>
      <td align=right width="15%" class="clefttitle">�����ˣ�</td>
      <td width="85%">
        <Input size=64 value="<%=KS.Setting(0)%>" name=sendername class="textbox"> </td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%" class="clefttitle">������Email��</td>
      <td width="85%">
        <Input size=64 value="<%=KS.Setting(11)%>" name=senderemail class="textbox"> </td>
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">�ʼ����ȼ���</td>
      <td>
  <Input name=Priority type=radio value=1> 
  �� 
  <Input type=radio value=3 name=Priority checked="checked"> 
  ��ͨ 
        <Input type=radio value=5 name=Priority> �� </td>
    </tr>
    <tr class=tdbg>
      <td style="text-align:center" colSpan=2>
  <Input id=Action type=hidden value=Send name=Action> 
  <Input id=Submit class="button" type=submit value=" �� �� " name=Submit>&nbsp; 
        <Input id=Reset class="button" type=reset value=" �� �� " name=Reset> </td>
    </tr>
</FORM>
  </table>
		<%
		end sub
		
		Sub Send()
			Server.ScriptTimeout=99999
			Dim InceptType
			InceptType = Trim(Request.Form("InceptType"))
			Title	 = Trim(Request.Form("title"))
			If KS.S("sendtype")="1" Then
			  Content  = Request.Form("Content")
			Else
			      Content  = Request.Form("mailtemplate")
				  Content  = Replace(Content,"{$GetSiteName}",KS.Setting(0))
				  Content  = Replace(Content,"{$GetSiteUrl}",KS.GetDomain)
				  Content  = Replace(Content,"{$SendDate}",Now)
				  GetMailContent
				  If Not IsArray(BodySQL) Then
					KS.Showerror("ϵͳ�Ҳ���" & request("sendday") & "���ڿɷ��͵�����!")
					Exit Sub
				  End If
	  			 If InceptType<>"-1" Then 
				  GetMailBody ""
				  Content=Replace(Content,"{$MailContent}",MailBody)
                 End If
			End If
			sendername =KS.G("sendername")
			senderemail=KS.G("senderemail")
			sendfile=Request.Form("sendfile")
			
			If Title="" or Content="" Then
				KS.Showerror("����д�ʼ������������!")
				Exit Sub
			End If
			Numc=0
			Select Case InceptType
			Case "-1" : SaveMsg_100() '�������û�����
			Case "0" : SaveMsg_0()	'�������û�
			Case "1" : SaveMsg_1()	'��ָ���û���
			Case "2" : SaveMsg_2()	'��ָ���û�
			Case "3" : SaveMsg_3()  'ָ������
			Case Else
				KS.Showerror("���������ŵ��û�!") : Exit Sub
			End Select
			Call KS.Alert("�����ɹ������η���"&Numc&"���û����������Ĳ�����","KS.UserMail.asp")
		End Sub
		
		Sub GetMailContent()
		  Dim Rs,SendNum,SendDay,Param,sql
		  SendNum=KS.ChkClng(KS.S("SendNum"))
		  SendDay=KS.ChkClng(KS.S("SendDay"))
		  If SendNum=0 Then SendNum=10
		  Param="Where I.Verific=1 And C.MailTF=1"
		  If SendDay<>0 Then
		    Param=Param & " and datediff(" & DataPart_D & ",i.adddate," & SQLNowString &")<" & SendDay
		  End If
		  
		  sql="Select top " & SendNum & " i.id,i.title,i.adddate,i.tid,I.ChannelID,i.InfoID,I.Fname From KS_ItemInfo I Inner Join KS_Class C On I.Tid=C.ID "& Param & " Order By i.Id Desc"
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open sql,conn,1,1
		  If Not RS.Eof Then
		    BodySQL=RS.GetRows(-1)
		  End If
		  RS.Close
		  Set RS=Nothing
		End Sub
		Sub GetMailBody(ClassID)
		  Dim Str,I,num
		  num=0
		  For I=0 To Ubound(BodySQL,2)
		    If KS.IsNul(ClassID) Or KS.FoundInArr(ClassID,BodySQL(3,I),",") Then
			num=num+1
		    Str=str &"<tr style=""background:#ffffff""><td style=""height:22px;text-align:center"">" & (num) & "��</td><td style=""text-align:center""><a href='" & KS.GetFolderPath(BodySQL(3,I)) &"' target='_blank'>" & KS.C_C(BodySQL(3,I),1) & "</a></td><td><a href='" & KS.GetItemURL(BodySQL(4,I),BodySQL(3,I),BodySQL(5,i),BodySQL(6,I)) &"' target='_blank'>" & BodySQL(1,I) & "</a></td><td>" & BodySQL(2,i) & "</td></tr>"
			End If
		  Next
		  If Not KS.IsNul(Str)  Then
		    Str="<table cellspacing=""1"" cellpadding=""1"" style=""background:#f1f1f1""><tr><td style=""height:28px;text-align:center;font-weight:bold;background:#cccccc"">���</td><td style=""text-align:center;font-weight:bold;background:#cccccc"">��Ŀ����</td><td style=""text-align:center;font-weight:bold;background:#cccccc"">����</td><td style=""text-align:center;font-weight:bold;background:#cccccc"">����ʱ��</td></tr>" & str 
			str=str & "</table>"
		  End If
		  MailBody=Str
		End Sub
		
		'�������û�����
		Sub SaveMsg_100()
			Dim Rs,Sql,i,ReturnInfo
			Sql = "Select Email,ID,ActiveCode,ClassID From KS_UserMail Where ActiveTF=1 Order By ID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				For i=0 To Ubound(SQL,2)
				  If KS.S("sendtype")="1" Then
			       Content  = Request.Form("Content")
				   ReturnInfo=SendMail(KS.Setting(12),  KS.Setting(13), KS.Setting(14), Title, SQL(0,i),sendername, Content,senderemail)
					IF ReturnInfo="OK" Then  Numc=Numc+1

			     Else
					  Call GetMailBody(SQL(3,I))  '����û�����Ȥ����Ŀ����
					  If Not KS.IsNul(MailBody) Then
						  Content=Replace(Content,"{$MailContent}",MailBody)
						  if Not KS.ISNul(SQL(0,i)) then
							 ReturnInfo=SendMail(KS.Setting(12),  KS.Setting(13), KS.Setting(14), Title, SQL(0,i),sendername, Replace(Content,"{$GetCancelUrl}",KS.GetDomain & "plus/mailsub.asp?action=del&id=" & sql(1,i) & "&activecode=" & sql(2,i)),senderemail)
							  IF ReturnInfo="OK" Then  Numc=Numc+1
						  end if
					  End If
				 End If
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'�������û�����
		Sub SaveMsg_0()
			Dim Rs,Sql,i
			Sql = "Select Email From KS_User Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				For i=0 To Ubound(SQL,2)
				  if Not IsNull(SQL(0,i)) and SQL(0,i)<>"" then
				     Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12),  KS.Setting(13), KS.Setting(14), Title, SQL(0,i),sendername, Content,senderemail)
					  IF ReturnInfo="OK" Then  Numc=Numc+1
				  end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'ָ���û���
		Sub SaveMsg_1()
		    GroupID = Replace(Request.Form("GroupID"),chr(32),"")
			If GroupID<>"" and Not Isnumeric(Replace(GroupID,",","")) Then
				ErrMsg = "����ȷѡȡ��Ӧ���û��顣"
			Else
				GroupID = KS.R(GroupID)
			End If
			Dim Rs,Sql,i
			Sql = "Select Email From KS_User Where GroupID in(" & GroupID & ") Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				For i=0 To Ubound(SQL,2)
				  if Not IsNull(SQL(0,i)) and SQL(0,i)<>"" then
				     Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), Title, SQL(0,i),sendername, Content,senderemail)
					  IF ReturnInfo="OK" Then  Numc=Numc+1
				  end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'��ָ���û�
		Sub SaveMsg_2()
			Dim inceptUser,Rs,Sql,i
			inceptUser = Trim(Request.Form("inceptUser"))
			If inceptUser = "" Then
				KS.Showerror("����дĿ���û�����ע�����ִ�Сд��")
				Exit Sub
			End If
			inceptUser = Replace(inceptUser,"'","")
			inceptUser = Split(inceptUser,",")
			For i=0 To ubound(inceptUser)
				SQL = "Select Email From KS_User Where UserName = '"&inceptUser(i)&"'"
				Set Rs = Conn.Execute(SQL)
				If Not Rs.eof Then
				  if Not IsNull(rs(0)) and rs(0)<>"" then
				     Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12),  KS.Setting(13), KS.Setting(14), Title, rs(0),sendername, Content,senderemail)
					  IF ReturnInfo="OK" Then  Numc=Numc+1
				  end if
				End If
			Next
			Rs.Close : Set Rs = Nothing
		End Sub
		'��ָ������
		Sub SaveMsg_3()
			Dim InceptEmail,Rs,Sql,i
			InceptEmail = Trim(Request.Form("InceptEmail"))
			If InceptEmail = "" Then
				KS.Showerror("����д�����͵��ʼ���ַ!")
				Exit Sub
			End If
			InceptEmail = Replace(InceptEmail,"'","")
			InceptEmail = Split(InceptEmail,",")
			For i=0 To ubound(InceptEmail)
				Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), Title, InceptEmail(i),sendername, Content,senderemail)
				IF ReturnInfo="OK" Then  Numc=Numc+1
			Next
		End Sub
        
		'�����ʼ�
		Sub MailOut()
		%>
		<br>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
<FORM action="?Action=DoExport" method=post>
    <tr class=title>
      <td class=title align=middle colSpan=2 height=22><B>�ʼ��б��������������ݿ�</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="24%" height=80>�����ʼ��б����ݿ⣺</td>
      <td width="76%" height=80>
  <Input id=ExportType type=hidden value=1 name=ExportType> &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp; 
<Select id=GroupID name=GroupID>
  <Option value=0 selected>ȫ����Ա</Option>
<%=KS.GetUserGroup_Option(0)%>
</Select> &nbsp;<font color=blue>��</font>&nbsp; 
  <Input id=ExportFileName maxLength=200 size=30 value=<%=KS.Setting(3)%>usermail.mdb name=ExportFileName> 
        <Input type=submit value=��ʼ name=Submit> </td>
    </tr>
</FORM>
  </table>
<br>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
<FORM action="?Action=DoExport" method=post>
    <tr class=title>
      <td class=title align=middle colSpan=2 height=22><B>�ʼ��б������������ı�</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="24%" height=80>�����ʼ��б��ı���</td>
      <td width="76%" height=80>
  <Input id=ExportType type=hidden value=2 name=ExportType> &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp; 
<Select id=GroupID name=GroupID>
  <Option value=0 selected>ȫ����Ա</Option>
<%=KS.GetUserGroup_Option(0)%>
</Select> 
</Select>&nbsp;<font color=blue>��</font>&nbsp; 
  <Input id=ExportFileName maxLength=200 size=30 value=<%=KS.Setting(3)%>usermail.txt name=ExportFileName> 
        <Input type=submit value=��ʼ���� name=Submit2> </td>
    </tr>
</FORM>
  </table>

		<%
		End Sub
		
		'�������ı��ļ�
		Sub DoExport()
		 Dim ExportFileName:ExportFileName=KS.G("ExportFileName")
		 Dim GroupID:GroupID=KS.G("GroupID")
		 Dim ExportType:ExportType=KS.G("ExportType")
		 Dim rs:set rs=KS.InitialObject("adodb.recordset")
		 Dim sqlstr,MailList,n
		   n=0
		  if GroupID="0" then
		    sqlstr="select email from ks_user"
		  else
		    sqlstr="select email from ks_user where groupid=" & groupid
		  end if
			 If ExportType=2 Then
			    		 rs.open sqlstr,conn,1,1
						 if not rs.eof then
						   do while not rs.eof
						      if rs(0)<>"" and not isnull(rs(0)) then
							  n=n+1
							  MailList=MailList & rs(0) & vbcrlf
							  end if
							  rs.movenext
						   loop
						 end if
						  rs.close:set rs=nothing
				Dim FSO:Set FSO = KS.InitialObject(KS.Setting(99))
				Dim FileObj:Set FileObj = FSO.CreateTextFile(Server.MapPath(ExportFileName), True) '�����ļ�
				FileObj.Write MailList
				FileObj.Close     '�ͷŶ���
				Set FileObj = Nothing:Set FSO = Nothing
			 Else
			      on error resume next
			     if CreateDatabase(ExportFileName)=true then
						Dim DataConn:Set DataConn = KS.InitialObject("ADODB.Connection")
	                    DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ExportFileName)
						If not Err Then
						   If Checktable("UserEmail",DataConn)=true Then
						     DataConn.Execute("drop table useremail")
						   end if
				             Dataconn.execute("CREATE TABLE [UserEMail] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,[Email] varchar(255) Not Null)")
						  rs.open sqlstr,conn,1,1
						 if not rs.eof then
						   do while not rs.eof
						      if rs(0)<>"" and not isnull(rs(0)) then
							  n=n+1
						      DataConn.Execute("Insert Into UserEmail(email) values('" & rs(0) &"')")
							  end if
							  rs.movenext
						   loop
						 end if
                          rs.close:set rs=nothing
						End if
						DataConn.Close:Set DataConn=Nothing
				 end if
			 
			 End If
		  response.write "<br><br><br><div align=center>�������!�ɹ������� <font color=red>" & n & "</font> ���ʼ���ַ��<a href=" & ExportFileName & ">������������</a>(�Ҽ�Ŀ�����Ϊ)  </div><br><br><br><br><br><br><br>"
		End Sub
		Function CreateDatabase(dbname)
		      if KS.CheckFile(dbname) then CreateDatabase=true:exit function
				dim objcreate :set objcreate=KS.InitialObject("adox.catalog") 
				if err.number<>0 then 
					set objcreate=nothing 
					CreateDatabase=false
					exit function 
				end if 
				'�������ݿ� 
				objcreate.create("data source="+server.mappath(dbname)+";provider=microsoft.jet.oledb.4.0") 
				if err.number<>0 then 
					CreateDatabase=false
					set objcreate=nothing 
					exit function
				end if 
				CreateDatabase=true
		End Function
		'������ݱ��Ƿ����	
		Function Checktable(TableName,DataConn)
			On Error Resume Next
			DataConn.Execute("select * From " & TableName)
			If Err.Number <> 0 Then
				Err.Clear()
				Checktable = False
			Else
				Checktable = True
			End If
		End Function
		
		Public Function SendMail(MailAddress, LoginName, LoginPass, Subject, Email, Sender, Content, Fromer)
	   'On Error Resume Next
		Dim JMail
		  Set jmail = Server.CreateObject("JMAIL.Message") '���������ʼ��Ķ���
			jmail.silent = true '����������󣬷���FALSE��TRUE��ֵj
			jmail.Charset = "GB2312" '�ʼ������ֱ���Ϊ����
			If sendfile="" Then
			' jmail.ContentType = "text/html" '�ʼ��ĸ�ʽΪHTML��ʽ,��������ʱ�ſ���
			End If
			jmail.AddRecipient Email '�ʼ��ռ��˵ĵ�ַ
			jmail.From = Fromer '�����˵�E-MAIL��ַ
			jmail.FromName = Sender
			  If LoginName <> "" And LoginPass <> "" Then
				JMail.MailServerUserName = LoginName '�����ʼ���������¼��
				JMail.MailServerPassword = LoginPass '��¼����
			  End If
			jmail.Subject = Subject '�ʼ��ı��� 
			JMail.Body = Content
			JMail.HTMLBody = Content
			Dim I,sendfileArr:SendFileArr=Split(sendfile,",")
			For I=0 To UBound(SendFileArr)
			 if trim(sendfileArr(i))<>"" Then
			  jmail.AddAttachment server.MapPath(trim(sendfileArr(i)))
			 End If
			Next
			JMail.Priority = 1'�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
			jmail.Send(MailAddress) 'ִ���ʼ����ͣ�ͨ���ʼ���������ַ��
			jmail.Close() '�رն���
		Set JMail = Nothing
		If Err Then
			SendMail = Err.Description
			Err.Clear
		Else
			SendMail = "OK"
		End If
	  End Function
End Class
%> 
