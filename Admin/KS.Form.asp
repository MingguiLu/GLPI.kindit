<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Form
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Form
        Private KS,KSCls,I
		Private MaxPerPage,CurrentPage,TotalPut,ID,RS
		Private IConnStr,IConn
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		  With Response
		  
		   If Not KS.ReturnPowerResult(0, "KSMS10006") Then          '���Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
		   End If
		   If KS.G("Action")="createtemplate" Then
		     Call AutoTemplate()
			 response.end
		   ElseIf KS.G("Action")="export" Then
		     Call export()
			 response.End()
		   End If
		    .Write "<html>"
			.Write "<title>ģ�ͻ�����������</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script src=""../ks_inc/kesion.box.js"" language=""JavaScript""></script>"
			.Write "</head>"
			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
		  If KS.G("Action")="replay" Then
		    Call Replay()
			Response.End()
		  End If
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""location.href='KS.Form.asp?action=Add';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=ϵͳ���� >> <font color=red>ϵͳ�Զ����</font>';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��ӱ�</span></li>"
			.Write "<li class='parent' onclick='location.href=""KS.Form.asp?action=total""'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>���ô���</span></li>"
             If KS.G("Action")="" Then
			.Write "<li class='parent' disabled"
		     Else
			.Write "<li class='parent'"
			 End If
			.Write " onclick='location.href=""KS.Form.asp"";'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>������ҳ</span></li>"
			.Write "</ul>"

		  Select Case KS.G("Action")
		   Case "result"  Call SubmitResult()
		   Case "resulthp"  Call SubmitResultHP()
		   Case "setstatus" Call setstatus()
		   Case "delinfo"  Call DelInfo()
		   Case "SetFormParam" Call SetFormParam() 
		   Case "Edit","Add"  Call FormManage()
		   Case "EditSave" Call FormSave()
		   Case "Del" Call FormDel()
		   Case "total" Call Total()
		   Case "template" Call FormTemplate()
		   Case "TemplateSave" Call TemplateSave()
		   Case "view" Call FormView()
		   Case "replaysave" Call ReplaySave()
		   case "Import" Import()
		   Case "ImportNext" importNext()
		   Case "ImportNext2" importNext2()
		   Case "modifyinfo" modifyinfo()
		   Case "DoResultSave" DoResultSave()
		   Case Else Call Main()
		  End Select
		  End With
		End Sub
 
		Sub Main()
		   With Response
			.Write "<script>"
			.Write "function document.onreadystatechange(){"
			.Write "parent.frames['BottomFrame'].Button1.disabled=true;"
			.Write "parent.frames['BottomFrame'].Button2.disabled=true;"
			.Write "}</script>"
			.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_Form Order By ID",conn,1,1
		    .Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write "  <td width='50' align=center>ID</td><td align=center>������</td><td align=center>��Ч��</td><td align=center>��¼</td><td align=center>״̬</td><td align=center>���������</td>"
			.Write "</tr>"
		  Do While Not RS.Eof 
		    .Write "<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			.Write "<td align=center class='splittd'>" & RS("ID")&"</td>"
			.Write "<td align=center class='splittd'>" & RS("FormName") &"</td>"
			.Write "<td align='center' class='splittd'>" & RS("StartDate") & "<br/>��<br/>" & RS("ExpiredDate") & "</td>"
			.Write "<td align=center class='splittd'><font color=red>" & conn.execute("select count(*) from " & rs("tablename"))(0) & "</font> ��</td>"
			.Write "<td align=center  class='splittd'>" 
			  If RS("Status")="1" Then .Write "����" Else .Write "<font color=red>����</font>"
			.Write "</td>"
			.Write "<td width='330' class='splittd'>"
			.Write "<strong>��Ŀ����:</strong> <a href='#' onClick=""SelectObjItem1(this,'�Զ���� >> <font color=red>���ֶι���</font>','Disabled','KS.FormField.asp?ItemID=" & rs("ID") & "');"">�ֶι���</a>��"
			.Write "<a href='#' onClick=""SelectObjItem1(this,'�Զ���� >> <font color=red>����ģ��</font>','gosave','KS.Form.asp?ItemID=" & rs("ID") & "&action=template');"">����ģ��</a>��"
			.Write "<a href='KS.Form.asp?ItemID=" & rs("ID") & "&action=view'>Ԥ��</a>��"
			
			

			.Write "<a href='?action=Edit&ID=" & rs("ID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=GoSave&OpStr=��ϵͳ >> <font color=red>ϵͳ�Զ����</font>';"">�޸�</a>��"
			 .Write "<a href='?action=Del&ID=" & rs("ID") & "' onclick='return(confirm(""�˲��������棬ȷ��ɾ����""))'>ɾ��</a>��"
			 			 
			 If RS("Status")="1" Then .Write "<a href='?Action=SetFormParam&Flag=FormOpenOrClose&ID=" & RS("ID") & "'>����</a>" Else .Write "<a href='?Action=SetFormParam&Flag=FormOpenOrClose&ID=" & RS("ID") & "'>����</a>"
			.Write "<br/><strong>��¼����:</strong> <a href='#' onClick=""SelectObjItem1(this,'�Զ���� >> <font color=red>�鿴�ύ���</font>','Disabled','KS.Form.asp?ItemID=" & rs("ID") & "&action=resulthp');"">���Ų鿴</a>��<a href='#' onClick=""SelectObjItem1(this,'�Զ���� >> <font color=red>�鿴�ύ���</font>','Disabled','KS.Form.asp?ItemID=" & rs("ID") & "&action=result');"">���Ų鿴</a> ��<a href='#' onClick=""SelectObjItem1(this,'�Զ���� >> <font color=red>�鿴�ύ���</font>','Disabled','KS.Form.asp?action=modifyinfo&FormID=" & rs("ID") & "');"">��Ӽ�¼</a>��<a href='?Action=Import&ID=" & RS("ID") & "'>��������</a></td></tr>"
			RS.MoveNext 
		  Loop
		    .Write "</table>"
			.Write "</div>"
		   RS.Close:Set RS=Nothing
		    .Write "</body>"
			.Write "</html>"
		  End With
		End Sub
		
		Sub FormDel()
		  on error resume next
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  Conn.BeginTrans
		  Dim TableName:TableName=Conn.Execute("select tablename from ks_form where id=" & ID)(0)
		  Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1016 and infoid=" & ID)
		  Conn.Execute("Drop Table " & TableName)
		  Conn.Execute("Delete From KS_Form Where ID=" & ID)
		  Conn.Execute("Delete From KS_FormField Where ItemID=" & ID)
		  If Err<>0 Then
		   Conn.RollBackTrans
		  Else
		   Conn.CommitTrans
		  End If
		  Response.Write "<script>alert('����Ŀɾ���ɹ�!');location.href='KS.Form.asp';</script>" 
		End Sub
        		
		Sub Total()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Form Where Status=1 order by ID asc",conn,1,1
		   With Response
		  	.Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write " <td align=center colspan=6>������Ŀ��ǰ̨���ô���</td>"
			.Write "</tr>"

		  Do While Not RS.Eof
			.Write "<tr height='25' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			
			If RS("PostByStep")="1" or not conn.execute("select top 1 FieldType From KS_FormField Where ItemID=" & RS("ID") & " And (FieldType=10 or FieldType=11)").eof Then
			.Write "<td width='50' class='splittd'></td><td width='140' class='splittd'><img src='images/37.gif'>&nbsp;<b>" & RS("FormName") & "</b></td><td class='splittd'><input type='text' value='&lt;iframe src=&quot;" & KS.Setting(2) & "/plus/form.asp?id=" & rs("id") & "&quot; width=&quot;550&quot; height=&quot;350&quot; allowtransparency=&quot;true&quot; frameborder=&quot;0&quot;&gt;&lt;/iframe&gt;' name='s" & rs(0) & "' size=60></td><td class='splittd'><input class=""button"" onClick=""jm_cc('s" & rs(0) & "')"" type=""button"" value=""���Ƶ�������"" name=""button""></td><td class='splittd'></td>"
			Else
			.Write "<td width='50'></td><td width='140'><img src='images/37.gif'>&nbsp;<b>" & RS("FormName") & "</b></td><td><input type='text' value='&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & KS.Setting(2) & "/plus/form.asp?id=" & rs("id") & "&quot;&gt;&lt;/script&gt;' name='s" & rs(0) & "' size=60></td><td><input class=""button"" onClick=""jm_cc('s" & rs(0) & "')"" type=""button"" value=""���Ƶ�������"" name=""button""></td><td></td>"
			End If
			
			.Write "</tr>"
		    RS.MoveNext
		  Loop
		   .Write "</table>"
		  End With
		  RS.Close:Set RS=Nothing
		  %>
		  <div style="color:red;padding-left:30px;margin-top:20px">
		   ����˵����ǰ̨ģ������ֻ�ǵ��������ұ����������ֶκͱ༭���ֶ�ʱ����<scrpit>����,�������iframe����,�����iframe���õ����ʵ�����iframe�Ŀ�͸�!
		  </div>
		   <script>
			function jm_cc(ob)
			{
				var obj=MM_findObj(ob); 
				if (obj) 
				{
					obj.select();js=obj.createTextRange();js.execCommand("Copy");}
					alert('���Ƴɹ���ճ������Ҫ���õ�html�����Ｔ��!');
				}
				function MM_findObj(n, d) { //v4.0
			  var p,i,x;
			  if(!d) d=document;
			  if((p=n.indexOf("?"))>0&&parent.frames.length)
			   {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
			   }
			  if(!(x=d[n])&&d.all) x=d.all[n];
			  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && document.getElementById) x=document.getElementById(n); return x;
			}
  </script>
		  <%
		End Sub
		
		Sub SetFormParam()
		   With Response
			   Dim ID:ID=KS.ChkClng(KS.G("ID"))
			   If ID=0 Then .Redirect "?": Exit Sub
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From KS_Form Where ID=" & ID,Conn,1,3
			   If RS.Eof Then
				 RS.Close:Set RS=Nothing
				.Redirect "?": Exit Sub
			   End If
		     If KS.G("Flag")="FormOpenOrClose" Then
			   If RS("Status")=1 Then 
					RS("Status")=0 
			   Else 
			    RS("Status")=1
			   end if
			 End If
			 RS.Update
			 RS.Close:Set RS=Nothing
			 .Write "<script>location.href='?';</script>"
		   End With
		End Sub
		
		Sub FormManage()
		Dim TimeLimit,AllowGroupID,useronce,onlyuser,shownum,PostByStep,StepNum,ToUserEmail
		Dim TempStr,SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt,i
		Dim FormName,ExpiredDate,StartDate,Status,Descript,TableName,UpLoadDir
		

		Dim ID:ID = KS.ChkClng(KS.G("ID"))
	'	On Error Resume Next
	   If KS.G("Action")="Edit" Then
			SqlStr = "select top 1 * from KS_Form Where ID=" & ID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			Status = RS("Status")
			FormName     = RS("FormName")
			TableName    = Replace(RS("TableName"),"KS_Form_","")
			UpLoadDir    = RS("UpLoadDir")
			StartDate    = RS("StartDate")
			TimeLimit    = RS("TimeLimit")
			ExpiredDate  = RS("ExpiredDate")
			TimeLimit    = RS("TimeLimit")
            AllowGroupID = RS("AllowGroupID")
			Descript     = RS("Descript")
			useronce     = RS("useronce")
			onlyuser     = RS("onlyuser")
			shownum      = RS("shownum")
			PostByStep   = RS("PostByStep")
			StepNum      = RS("StepNum")
			ToUserEmail  = RS("ToUserEmail")
		Else
		      Status=1:TimeLimit = 0:StartDate=Now():ExpiredDate=Now()+10:AllowGroupID="":useronce=0:onlyuser=0:shownum=1:UpLoadDir="form":PostByStep=0:StepNum=1:ToUserEmail=0
		End If
		
		With Response
		.Write "<html>"&_
		"<title>��ӱ�</title>" &_
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" &_
		"<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"&_
		"<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"&_
		"<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & _
		"<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & _
		"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"&_
		"<body>" &_
		"<table width='100%' border='0' cellspacing='0' cellpadding='0'>"&_
		"  <tr>"&_
		"	<td height='25' class='sort'>�Զ��������</td>"&_
		" </tr>"&_
		" <tr><td height=5></td></tr>"&_
		"</table>" & _
			
		"<div class=tab-page id=Formpanel>"& _
		"<form name=""myform"" method=""post"" action=""KS.fORM.asp?Action=EditSave&ID=" & ID & """ onSubmit=""return(CheckForm())"">" & _
        " <SCRIPT type=text/javascript>"& _
        "   var tabPane1 = new WebFXTabPane( document.getElementById( ""Formpanel"" ), 1 )"& _
        " </SCRIPT>"& _
             
		" <div class=tab-page id=site-page>"& _
		"  <H2 class=tab>������Ϣ</H2>"& _
		"	<SCRIPT type=text/javascript>"& _
		"				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"& _
		"	</SCRIPT>" & _
		"<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'> <div align=""right""><strong>��״̬��</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""Status"" value=""1"" "
		If Status = 1 Then .Write (" checked")
		.Write ">"
		.Write "����"
		.Write "  <input type=""radio"" name=""Status"" value=""0"" "
		If Status = 0 Then .Write (" checked")
		.Write ">"
		.Write "�ر�</td>"
		.Write "    </tr>"

%>
		<script>
		 function CheckForm()
		 {
		  if ($("input[name=FormName]").val()=="")
		  {
		   $("input[name=FormName]").focus();
		   alert('�����������');
		   return false;
		  }
		  
		  $("form[name=myform]").submit();
		 }
		 
		 function changedate()
		 {
		   val=$("input[name=TimeLimit][checked=true]").val();
		   if (val==1){
		    $("#BeginDate").show()
		    $("#EndDate").show();		
		   }
		   else{
		    $("#BeginDate").hide();
		    $("#EndDate").hide();		
		   }
		 }
		 function changepage()
		 {
		   val=$("input[name=PostByStep][checked=true]").val();
		   if (val==1){
		    $("#StepNumArea").show();
		   }
		   else{
		    $("#StepNumArea").hide();
		   }
		 }
	
		</script>

		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>�����ƣ�</strong></div></td>      
			<td height="30"> <input name="FormName" class="textbox" type="text" value="<%=FormName%>" size="30"> �磺�����������</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><strong>���ݱ����ƣ�</strong></td>      
			<td height="30"> KS_Form_<input name="TableName"<%If KS.G("Action")="Edit" then response.write " disabled"%> size="14" class="textbox" type="text" value="<%=TableName%>" size="30"> 
			<br><font color=red>˵�����������ݱ���޷��޸ģ������û����������ݱ���"KS_Form_"��ͷ</font></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><strong>�ϴ�Ŀ¼��</strong></td>
			<td><%=KS.Setting(91)%><input name="UpLoadDir" size="14" class="textbox" type="text" value="<%=UpLoadDir%>" size="30"> 
			<br><font color=blue>˵����ֻ������ĸ�����ֵ���ϣ��ұ�����/������</font></td> 
		</tr>
 
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>���÷ֲ��ύ����</strong></div></td>      
			<td height="30"> <input onClick="changepage()" name="PostByStep" type="radio" value="1"<%if PostByStep="1" Then Response.Write " Checked"%>>���� <input onClick="changepage()" type="radio" value="0" name="PostByStep"<%if PostByStep="0" Then Response.Write " Checked"%>>������
			<br/><font color=blue>����Ҫ�ռ����û����Ͻ϶�ʱ,�������÷ֲ��ύ����</font>
			</td> 
		</tr>
		<tr id="StepNumArea" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>�ֲ����ã�</strong></div></td>      
			<td height="30"> �û���Ϊ<input name="StepNum" size="4" class="textbox" type="text" value="<%=StepNum%>" style="text-align:center">���ύ</td> 
		</tr>

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>����ע��</strong></div></td>      
			<td height="30"> <textarea name="Descript" class="textbox" style="width:400px;height:90px"><%=Descript%></textarea></td> 
		</tr>
		</table>
		</div>
		 <div class=tab-page id="formset">
		  <H2 class=tab>ѡ������</H2>
			<SCRIPT type=text/javascript>
				 tabPane1.addTabPage( document.getElementById( "formset" ) );
			</SCRIPT>
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>����ʱ�����ƣ�</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""1"" "
		If TimeLimit = 1 Then .Write (" checked")
		.Write ">"
		.Write "����"
		.Write "  <input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""0"" "
		If TimeLimit = 0 Then .Write (" checked")
		.Write ">"
		.Write "������"
		
			%>
			</td> 
		</tr>

		<tr ID="BeginDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">     
		<td height="30" class="clefttitle"align="right"><div><strong>��Чʱ�䣺</strong></div></td>     
		<td height="30"><input name="StartDate" id='StartDate' class="textbox" type="text" value="<%=StartDate%>" size="24"><br><font color=#ff0000>���ڸ�ʽ��0000-00-00 00:00:00</font></td>   
		</tr> 
		
		<tr ID="EndDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>ʧЧʱ�䣺</strong></div></td>      
			<td height="30"> <input name="ExpiredDate" id="ExpiredDate" class="textbox" type="text" value="<%=ExpiredDate%>" size="30"><br><font color=#ff0000>���ڸ�ʽ��0000-00-00 00:00:00</font></td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>ֻ�����Ա�ύ��</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""onlyuser"" value=""1"" "
		If onlyuser = 1 Then .Write (" checked")
		.Write ">"
		.Write "��"
		.Write "  <input type=""radio"" name=""onlyuser"" value=""0"" "
		If onlyuser = 0 Then .Write (" checked")
		.Write ">"
		.Write "����"
		
			%>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>ÿ����Աֻ�����ύһ�Σ�</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""useronce"" value=""1"" "
		If useronce = 1 Then .Write (" checked")
		.Write ">"
		.Write "��"
		.Write "  <input type=""radio"" name=""useronce"" value=""0"" "
		If useronce = 0 Then .Write (" checked")
		.Write ">"
		.Write "����"
		
			%>
			</td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>���ύ������͵����䣺</strong></div>
			<font color=red>��Ҫ���û���д����ʱ��������ô˹��ܽ��Զ����û����ύ��������û�������͹���Ա����</font>
			</td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""ToUserEmail"" value=""1"" "
		If ToUserEmail = 1 Then .Write (" checked")
		.Write ">"
		.Write "����"
		.Write "  <input type=""radio"" name=""ToUserEmail"" value=""0"" "
		If ToUserEmail = 0 Then .Write (" checked")
		.Write ">"
		.Write "������"
		
			%>
			</td> 
		</tr>

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>�û������ƣ�</strong></div><font color=#ff0000>�����ƣ��벻Ҫѡ</font></td>      
			<td height="30"><%=KS.GetUserGroup_CheckBox("AllowGroupID",AllowGroupID,5)%> </td> 
		</tr>
				<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>��ʾ��֤�룺</strong></div></td>      
			<td height="30"> 
			
			<%
			response.write "<input type=""radio"" name=""shownum"" value=""1"" "
			If shownum = 1 Then .Write (" checked")
			.Write ">"
			.Write "��ʾ"
			.Write "  <input type=""radio"" name=""shownum"" value=""0"" "
			If shownum = 0 Then .Write (" checked")
			.Write ">"
			.Write "����ʾ"
		
			%>
			</td> 
	    	</tr>
			</table>
        </div>
		<script>changedate();changepage();</script>
		<%
		.Write "</form>"
		.Write "</div>"
		End With
		End Sub
		
		'��ģ�����
		Sub FormTemplate()
		 Dim FormID:FormID=KS.ChkClng(KS.G("ItemID"))
		 Dim RS,Template,FormName,PostByStep,StepNum,Step,K
		 
		 Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select FormName,PostByStep,StepNum,Template From KS_Form Where ID=" & FormID,conn,1,1
		 If RS.EOF And RS.Bof Then
		  Response.Write "<script>alert('error!');history.back();</script>"
		  Exit Sub
		 Else
		   FormName=RS(0):PostByStep=RS(1):StepNum=RS(2):Template=RS(3)
		 End If
		 RS.Close
         If Template="" Or IsNull(Template) Then Template=" "
		 Template=Split(Template,"$aaa$")
		%>
		<html>
		<title>��ģ�����</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		<script src="../KS_Inc/jquery.js" language="JavaScript"></script>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
         <script language = 'JavaScript'>
				   
					function LoadTemplate()
					{   
					   if ($("#autocreate").attr("checked")==true)
					    { 
							var url='KS.Form.asp';
							$.ajax({
								  url: url,
								  cache: false,
								  data: "action=createtemplate&formid="+$("#FormID").val(),
								  success: function(s){
									s=s.split("$aaa$");
								   <%For K=1 To StepNum%>
									  $('textarea[name=Content<%=K%>]').val(s[<%=K-1%>]);
									  if ($('textarea[name=Content<%=K%>]').val()=='undefined')
									  $('textarea[Content<%=K%>]').val('����ӱ���!');
								   <%Next%>
								  }
								});
							  
						}
						else
						{
						  $('#Content').val('');
						}
					}	

		            function show_ln(txt_ln,txt_main){
			            var txt_ln  = document.getElementById(txt_ln);
			            var txt_main  = document.getElementById(txt_main);
			            txt_ln.scrollTop = txt_main.scrollTop;
			            while(txt_ln.scrollTop != txt_main.scrollTop)
			            {
				            txt_ln.value += (i++) + '\n';
				            txt_ln.scrollTop = txt_main.scrollTop;
			            }
			            return;
		            }
		            function editTab(){
			            var code, sel, tmp, r
			            var tabs=''
			            event.returnValue = false
			            sel =event.srcElement.document.selection.createRange()
			            r = event.srcElement.createTextRange()
			            switch (event.keyCode){
				            case (8) :
				            if (!(sel.getClientRects().length > 1)){
					            event.returnValue = true
					            return
				            }
				            code = sel.text
				            tmp = sel.duplicate()
				            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
				            sel.setEndPoint('startToStart', tmp)
				            sel.text = sel.text.replace(/\t/gm, '')
				            code = code.replace(/\t/gm, '').replace(/\r\n/g, '\r')
				            r.findText(code)
				            r.select()
				            break
			            case (9) :
				            if (sel.getClientRects().length > 1){
					            code = sel.text
					            tmp = sel.duplicate()
					            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
					            sel.setEndPoint('startToStart', tmp)
					            sel.text = '\t'+sel.text.replace(/\r\n/g, '\r\t')
					            code = code.replace(/\r\n/g, '\r\t')
					            r.findText(code)
					            r.select()
				            }else{
					            sel.text = '\t'
					            sel.select()
				            }
				            break
			            case (13) :
				            tmp = sel.duplicate()
				            for (var i=0; tmp.text.match(/[\t]+/g) && i<tmp.text.match(/[\t]+/g)[0].length; i++) tabs += '\t'
				            sel.text = '\r\n'+tabs
				            sel.select()
				            break
			            default  :
				            event.returnValue = true
				            break
				            }
			            }
					 function CheckForm()
					 {
					 		  $("#myform").submit();
					 }
		            //-->
		            </script>

	  <body>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'>
		  <tr>
			<td height='25' class='sort'>�Զ����ģ�����</td>
		 </tr>
		 <tr><td height=5></td></tr>
		</table>
		<form name="myform" id="myform" action="KS.Form.asp?action=TemplateSave" method="post">
		<input type="hidden" value="<%=formid%>" name="FormID" id="FormID">
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		   <tr class='tdbg'>
		      <td class='clefttitle' width="120" align="right"><strong>�����ƣ�</strong></td>
		     <td height="30"> <font color=red><%=FormName%></font></td>
		   </tr>
		   <tr class='tdbg'>
		      <td class='clefttitle' width="120" align="right"><strong>�Զ�����ģ�壺</strong></td>
		     <td height="30">
			 <input type='checkbox' name='autocreate' id='autocreate' value='1' onClick="LoadTemplate()">�Զ�����
			 <font color=red>��ʾ����һ������ģ�壬���Ե���Զ����ɣ�</font>
			 </td>
		   </tr>
		  
		   <% 
		   on error resume next
		   For K=1 To StepNum%> 
		   <tr class='tdbg'>
		      <td class='clefttitle' align="right" width="130"><strong>��ģ��<%If PostByStep=1 Then %>(��<font color=red><%=K%></font>��)<%End If%>��</strong>
			  <%If K>1 Then Response.Write "<br><font color=red>�������{$HiddenFields}��ǩ</font>" %>
			  </td>
		     <td height="280">
			 <textarea id='txt_ln<%=K%>' name='rollContent' cols='6' style='overflow:hidden;height:280px;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly><%
		 Dim N
		 For N=1 To 3000
			Response.Write N & "&#13;&#10;"
		 Next
		 On Error Resume Next
		 %>
		 </textarea>
		 <textarea name='Content<%=K%>' style='width:600px;height:280px' ROWS='15' id='txt_main<%=K%>' onkeydown='editTab()' onscroll="show_ln('txt_ln<%=K%>','txt_main<%=K%>')" wrap='on'><%=server.HTMLEncode(Template(K-1))%></textarea>
			 </td>
		   </tr>
		   <%Next%>
		 </table>  
		 </form>
		<%
		
		End Sub
		
		Sub AutoTemplate()
		 Response.CharSet="gb2312" 
		 Dim ShowNum,PostByStep,StepNum,K,Param,S,KK
		 Dim SQL,N,O_Arr,O_Len,F_V,BrStr,O_Value,O_Text
		 Dim FormID:FormID=KS.ChkClng(KS.G("FormID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 ShowNum,PostByStep,StepNum From KS_Form Where ID=" & FormID,conn,1,1
		 If Not RS.Eof Then
		  ShowNum=RS(0):PostByStep=RS(1):StepNum=RS(2)
		 End If
		 RS.Close
		 
		 For S=1 To StepNum
		     SQL=""
		     Param="Where ShowOnForm=1 and ItemID=" & FormID 
			 If PostByStep=1 Then Param=Param & " and step=" & S
			 RS.Open "Select Title,FieldName,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,AllowFileExt,MaxFileSize,FieldID,ParentFieldName,ShowUnit,UnitOptions From KS_FormField " & Param & " order by orderid",conn,1,1
			 If Not RS.Eof Then SQL=RS.GetRows(-1)
			 RS.Close
			 If Not IsArray(SQL) Then Response.Write "�ñ���û����ӱ���!":Response.End
			 If PostByStep=1 Then
			 Response.Write "<div style=""text-align:center"">�� " & S & " ��</div>" & vbcrlf
			 End If
			 Response.Write "<table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""1"">" & vbcrlf
			 Response.Write "<form name=""myform"" action=""" & ks.setting(3) & "plus/form.asp"" method=""post""> " &vbcrlf
			 If (PostByStep=1 And S=StepNum)  Or PostByStep=0 Then
			 Response.Write "<input type=""hidden"" value=""Save"" name=""action"">" & vbcrlf
			 Else
			 Response.Write "<input type=""hidden"" value=""Next"" name=""action"">" & vbcrlf
			 End If
			 Response.Write "<input type=""hidden"" value=""" & FormID & """ name=""id"">" & vbcrlf
			 If PostByStep=1 Then
			 Response.Write "<input type=""hidden"" value=""" & S & """ name=""Step"">" & vbcrlf
			 End If
			 If S>1 Then	 Response.Write "{$HiddenFields}" & vbcrlf
			 
			 For K=0 To Ubound(SQL,2)
			 If SQL(12,K)="0" Or KS.IsNul(SQL(12,K)) Then
			 Response.Write " <tr class=""tdbg"">" & vbcrlf
			 Response.Write "  <td align=""right"" class=""lefttdbg"">" & SQL(0,K) & "��</td>" & vbcrlf
			 if KS.ChkClng(SQL(3,K))=10 Then
			 Response.Write "  <td style=""height:" & SQL(8,K) & "px;width:" & KS.ChkClng(SQL(7,K))+100 &"px;"">" 
			 Else
			 Response.Write "  <td>" 
			 End If
			 
			 Select Case SQL(3,K)
				Case 2
				  Response.Write "<textarea style=""width:" & SQL(7,K) & "px;height:" & SQL(8,K) &"px"" rows=""5"" name=""" & SQL(1,K) & """>" & SQL(4,K) & "</textarea>"
			   Case 3,11
			     If SQL(3,K)=11 Then
				  Response.Write "<select class=""upfile"" onchange=""fill" & SQL(1,K) &"(this.value)"" style=""width:" & SQL(7,K) & "px"" name=""" & SQL(1,K) & """><option value=''>---��ѡ��---</option>"
				 Else
				  Response.Write "<select class=""upfile"" style=""width:" & SQL(7,K) & "px"" name=""" & SQL(1,K) & """>"
				 End If
				  O_Arr=Split(SQL(5,K),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
					If O_Arr(N)<>"" Then
						F_V=Split(O_Arr(N),"|")
						If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						Else
							O_Value=F_V(0):O_Text=F_V(0)
						End If						   
						If SQL(4,K)=O_Value Then
							Response.Write "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
						Else
							Response.Write "<option value=""" & O_Value& """>" & O_Text & "</option>"
						End If
					End If
				  Next
				  Response.Write "</select>"
                  '�����˵�
					If SQL(3,K)=11  Then
						Dim JSStr
						Response.Write  GetLDMenuStr(FormID,SQL(1,k),JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
					End If				  
			  Case 6
				  O_Arr=Split(SQL(5,K),vbcrlf): O_Len=Ubound(O_Arr)
				  If O_Len>1 And Len(SQL(5,I))>50 Then BrStr="<br>" Else BrStr=""
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If SQL(4,K)=O_Value Then
						Response.Write "<input type=""radio"" name=""" & SQL(1,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
					Else
						Response.Write "<input type=""radio"" name=""" & SQL(1,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
					End If
				   End If
				  Next
			 Case 7
				   O_Arr=Split(SQL(5,K),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If KS.FoundInArr(SQL(4,K),O_Value,",")=true Then
						Response.Write "<input type=""checkbox"" name=""" & SQL(1,K) & """ value=""" & O_Value& """ checked>" & O_Text
					Else
						Response.Write "<input type=""checkbox"" name=""" & SQL(1,k) & """ value=""" & O_Value& """>" & O_Text
					End If
					End If
				  Next
			 Case 10
			        Response.Write "<script type=""text/javascript"" src=""" & KS.GetDomain& "editor/ckeditor.js"" mce_src=""" & KS.GetDomain& "editorckeditor.js""></script>"
					Response.Write "<textarea id=""" & SQL(1,K) &""" name=""" & SQL(1,K) &""">"& Server.HTMLEncode(SQL(4,K)) &"</textarea>"
					Response.Write "<script type=""text/javascript"">CKEDITOR.replace('" & SQL(1,K) &"', {width:""100%"",height:""" & SQL(8,K) & """,toolbar:""Basic"",filebrowserBrowseUrl :""" & KS.GetDomain& "editor/ksplus/SelectUpFiles.asp"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>"
					
		
			 Case Else
				Response.Write "<input type=""text"" class=""upfile"" style=""width:" & SQL(7,K) & "px"" name=""" & SQL(1,K) & """ value=""" & SQL(4,K) & """>"
			 End Select
			 
              If SQL(13,K)="1" Then 
					  Response.Write " <select name=""" & SQL(1,K) & "_Unit"" id=""" & SQL(1,K) & "_Unit"">"
					  If Not KS.IsNul(SQL(14,K)) Then
				       Dim UnitOptionsArr:UnitOptionsArr=Split(SQL(14,K),vbcrlf)
					   For KK=0 To Ubound(UnitOptionsArr)
					     response.write "<option value='" & UnitOptionsArr(kk) & "'>" & UnitOptionsArr(kk) & "</option>"                 
					   Next
					  End If
					  response.write "</select>"
			 End If
				   			
			 If SQL(6,K)=1 Then Response.Write "<font color=""red""> * </font>"
			 If SQL(2,K)<>"" Then Response.Write " <span style=""margin-top:5px"">" &  SQL(2,K) & "</span>"
			 If SQL(3,K)=9 Then Response.Write "���ϴ��ļ�����" & SQL(9,K) & ",��С" & SQL(10,K) & " KB<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='" &KS.Setting(3) & "user/User_UpFile.asp?Type=Field&FieldID=" & SQL(11,K) & "' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
			 Response.Write "  </td>" & vbcrlf
			 Response.Write "</tr>" & vbcrlf
			 End If
			 Next
			 IF ShowNum=1 And  S=StepNum Then
			 Response.Write "<tr class=""tdbg""><td class=""lefttdbg"" align=""right"">��֤�룺</td><td><input name=""Verifycode"" type=""text"" name=""textbox"" size=5><IMG style=""cursor:pointer"" src=""" & KS.Setting(3) & "plus/verifycode.asp"" onClick=""this.src='" &KS.Setting(3) & "plus/verifycode.asp?n='+ Math.random();"" align=""absmiddle""></td></tr>"  &vbcrlf
			 End If
			 If S=StepNum Then
			 Response.Write "<tr><td colspan=""2"" class=""subtdbg"" align=""center""><input type=""submit"" value=""ȷ���ύ"" name=""submit1""></td></tr>"  & vbcrlf
			 Else
			 Response.Write "<tr><td colspan=""2"" class=""subtdbg"" align=""center""><input type=""submit"" value=""OK����һ��"" name=""submit1""></td></tr>"  & vbcrlf
			 End If
			 Response.Write "</form>" &vbcrlf
			 Response.Write "</table>" & vbcrlf
			 Response.Write "$aaa$" & vbcrlf
			
		   Next	 
			 
			 
		End Sub
		
		'ȡ�������˵�
		   Function GetLDMenuStr(ItemID,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str
		     Dim RSL:Set RSL=Conn.Execute("Select Top 1 FieldName,Title,Options,Width From KS_FormField Where itemid=" & ItemID & " and ParentFieldName='" & ParentFieldName & "'")
			 If Not RSL.Eof Then
			     Str=Str & " <select name='" & RSL(0) & "' id='" & RSL(0) & "' onchange='fill" & RSL(0) & "(this.value)' style='width:" & RSL(3) & "px'><option value=''>--��ѡ��--</option>"
				 JSStr=JSStr & "var sub" &ParentFieldName & " = new Array();"
				  Options=RSL(2)
				  OArr=Split(Options,Vbcrlf)
				  For I=0 To Ubound(OArr)
				    Varr=Split(OArr(i),"|")
					If Ubound(Varr)=1 Then 
					 V=Varr(0):F=Varr(1)
					Else
					 V=trim(OArr(i))
					 F=trim(OArr(i))
					End If
				    JSStr=JSStr & "sub" & ParentFieldName&"[" & I & "]=new Array('" & V & "','" & F & "')" &vbcrlf
				  Next
				 Str=Str & "</select>"
				 JSStr=JSStr & "function fill"& ParentFieldName&"(v){" &vbcrlf &_
							   "$('#"& RSL(0)&"').empty();" &vbcrlf &_
							   "$('#"& RSL(0)&"').append('<option value="""">--��ѡ��--</option>');" &vbcrlf &_
							   "for (i=0; i<sub" &ParentFieldName&".length; i++){" & vbcrlf &_
							   " if (v==sub" &ParentFieldName&"[i][0]){document.getElementById('" & RSL(0) & "').options[document.getElementById('" & RSL(0) & "').length] = new Option(sub" &ParentFieldName&"[i][1], sub" &ParentFieldName&"[i][1]);}}" & vbcrlf &_
							   "}"

				 GetLDMenuStr=str & GetLDMenuStr(itemid,RSL(0),JSStr)
			 Else
			     JSStr=JSStr & "function fill" & ParentFieldName &"(v){}"				 
			 End If
			     
		   End Function
				
		'��ģ�����
		Sub FormView()
		 Dim FormID:FormID=KS.ChkClng(KS.G("ItemID"))
		 Dim PostByStep:PostByStep=LFCls.GetSingleFieldValue("Select PostByStep From KS_Form Where ID=" & FormID)
		%>
		<html>
		<title>Ԥ����</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
	  <body>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'>
		  <tr>
			<td height='25' class='sort'>�Զ����Ч��Ԥ��</td>
		 </tr>
		 <tr><td height=5></td></tr>
		</table>
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		   <tr class='tdbg'>
		      <td class='clefttitle' height="25" align="center"><strong>�����ƣ�<font color=red><%=Conn.Execute("Select FormName From KS_Form Where ID=" & FormID)(0)%></font></strong></td>
		   </tr>
		   <tr class='tdbg'>
		     <td>
			 <%If PostByStep=1 or not conn.execute("select top 1 FieldType From KS_FormField Where ItemID=" & FormID & " And (FieldType=10 or FieldType=11)").eof Then%>
			  <iframe src="../plus/form.asp?id=<%=formid%>" frameborder="0" width="550" height="500" allowtransparency="true"></iframe>
			 <%else%>
			 <script src="../plus/form.asp?id=<%=formid%>"></script>
			 <%end if%>
			 </td>
		   </tr>
		   <tr class='tdbg'>
		      <td class='clefttitle' height="25" align="center"><input type="button" class="button" onClick="SelectObjItem1(this,'�Զ���� >> <font color=red>ģ���޸�</font>','gosave','KS.Form.asp?ItemID=<%=FormID%>&action=template');" value="�޸�ģ��"></td>
		   </tr>
		 </table>  
		<%
		
		End Sub

		
		Sub FormSave()
		    Dim ExpiredDate,StartDate,I,OpName,ID:ID=KS.ChkClng(KS.G("ID"))
			StartDate=KS.G("StartDate")
			ExpiredDate=KS.G("ExpiredDate")
			If Not IsDate(StartDate) Then Call KS.AlertHistory("��Ч���ڸ�ʽ����ȷ",-1):response.end
			If Not IsDate(ExpiredDate) Then Call KS.AlertHistory("ʧЧ���ڸ�ʽ����ȷ",-1):response.end
			If ID=0 and Not Conn.Execute("select top 1 id from ks_form where tablename='KS_Form_" & KS.G("TableName") &"'").eof then Call KS.AlertHistory("���ݱ��Ѵ��ڣ�",-1):response.end
			on error resume next
			Conn.BeginTrans
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Form Where ID=" & ID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				RS("TableName")= "KS_Form_" & KS.G("TableName")
				OpName      = "���"
			Else
			    OpName="�޸�"
			End If
				RS("FormName")= KS.G("FormName")
				RS("UploadDir")= KS.G("UpLoadDir")
				RS("Status") = KS.G("Status")
				RS("TimeLimit")   = KS.ChkClng(KS.G("TimeLimit"))
				RS("StartDate")     = startdate
				RS("ExpiredDate")    = ExpiredDate
				RS("useronce") =KS.ChkClng(KS.G("useronce"))
				rs("onlyuser")=KS.ChkClng(KS.G("onlyuser"))
				rs("shownum")=ks.chkclng(ks.g("shownum"))
				RS("AllowGroupID")     = KS.G("AllowGroupID")
                RS("Descript")    = KS.G("Descript")
				RS("PostByStep")  = KS.ChkClng(KS.G("PostByStep"))
				RS("StepNum")     = KS.ChkClng(KS.G("StepNum"))
				RS("ToUserEmail") = KS.ChkClng(KS.G("ToUserEmail"))
				RS.Update
				RS.Close
				Set RS=Nothing
				
				If OpName="���" Then
				 Dim sql:sql="CREATE TABLE [KS_Form_" & KS.G("TableName") & "] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_KS_Form_" & KS.G("TableName") & " PRIMARY KEY,"&_
						"UserName nvarchar(100),"&_
						"UserIP nvarchar(100),"&_
						"AddDate datetime,"&_
						"[Note] text,"&_
						"Status tinyint default 0)"
				 Conn.Execute(sql)
				End If
				if err<>0 then
					Conn.RollBackTrans
					Call KS.AlertHistory("��������������" & replace(err.description,"'","\'"),-1):response.end
				else
					Conn.CommitTrans
					Response.Write ("<script>alert('" & OpName & "�Զ����ɹ�!');location.href='KS.Form.asp';</script>")
				end if
		End Sub
		
		Sub SubmitResult()
		ID=KS.ChkClng(KS.G("itemID"))
		Dim TableName,SQL,II
		TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & ID)
		MaxPerPage = 10     'ȡ��ÿҳ��ʾ����
		If KS.G("page") <> "" Then
			  CurrentPage = KS.ChkClng(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		 with response
		 %>
		  <script>
			function ShowReplay(formid,id)
			{  
			onscrolls=false;  //ȡ������
			new KesionPopup().PopupCenterIframe("�ظ�����¼","KS.Form.asp?Action=replay&formid="+formid+"&id=" +id,650,400,'no')
			 }
			</script>
			<div style="margin:8px"><a href="KS.Form.asp?ItemID=<%=id%>&action=resulthp"><strong>������ʾ��¼</strong></a> | <a href="KS.Form.asp?ItemID=<%=id%>&action=result"><strong>������ʾ��¼</strong></a></div>
		 <%
		    .Write ("<div sstyle=""height:94%; overflow: auto; width:100%"" align=""center"">")
		 	.Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write "  <td width='40' align='center'>ID��</td><td align=center>�ύ����</td><td align=center>���������</td>"
			.Write "</tr>"
			set rs=server.createobject("adodb.recordset")
			rs.open "select FieldName,title,MustFillTF,FieldType from ks_formfield where itemid=" & KS.ChkClng(KS.G("itemID")) & " and ShowOnForm=1 order by orderid",conn,1,1
			If Not RS.Eof Then SQL=RS.GetRows(-1)
			RS.Close
			rs.open "select * from " & TableName & " order by adddate desc" ,conn,1,1
			 If Not RS.EOF Then
					totalPut = Conn.Execute("Select count(1) From " & TableName)(0)
							If CurrentPage < 1 Then	CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage <> 1 Then
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								Else
									CurrentPage = 1
								End If
							End If
							  dim k,i:i=1
							  do while not rs.eof
							   response.write "<tr><td width=40 align='center'>" & rs("id") & "</td>"
							   response.write "<td align='left'>"
							   If IsArray(SQL) Then
								 response.write "<table width='100%' border='0'>"
								 For II=0 To Ubound(SQL,2)
								  response.write "<tr>"
								  response.write "<td width='100' align='right'><b>" & sql(1,ii) & "��</b></td>"
								  response.write "<td>" & rs(trim(sql(0,ii))) & "</td>"
								  response.write "</tr>"
								 Next
							   end if
							   response.write "</table>"
							   response.write "</td>"
							   response.write "<td>"
							   response.write "ʱ �䣺"  &rs("adddate") & "<br>IP��ַ��" & rs("userip") & "<br>�� ����" & rs("username")
							   response.write "<br>״ ̬��"
							   select case rs("status")
							   case 0
								response.write "<font color=red>δ��</font>"
							   case 1
								response.write "<font color=green>�Ѷ�</font>"
							   case 2
								response.write "<font color=#ff6600>����</font>"
							   case 3
								response.write "����"
							   end select
							   
							   if not isnull(rs("note")) and rs("note")<>"" then response.write "&nbsp;&nbsp;<a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");""><font color=blue>�ѻظ�</font></a>"
							   
							   response.write "<br>�� ����<a href=""?action=delinfo&FormID=" & ID&"&id=" & rs("Id") & """ onclick=""return(confirm('ȷ��ɾ����?'))"">ɾ��</a> <a href='KS.Form.asp?action=modifyinfo&FormID=" & id & "&id=" & rs("id") & "'>�޸�</a> <a href='?action=setstatus&v=1&FormID=" & ID&"&id=" & rs("id") & "' title='��Ϊ�Ѷ�'>�Ѷ�</a> <a href='?action=setstatus&v=2&FormID=" & ID&"&id=" & rs("id") & "' title='��Ϊ����'>����</a> <a href='?action=setstatus&v=3&FormID=" & ID&"&id=" & rs("id") & "' title='��Ϊ����'>����</a> <a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");"">�ظ�</a>"
							   response.write "</td>"
							   response.write "</tr>" 
							   Response.Write("<tr><td colspan=3><hr size=1 color=#cccccc></td></tr>")
							  rs.movenext
							  i=i+1
							  if i>maxperpage then exit do
							  loop

			 End If
			  .Write ("<tr> ")
			  .Write ("<td height=""30"" colspan=""3""><input type='button' class='button' onclick='window.print();' value='��ӡ��ҳ���¼'> <font color=red>��ܰ��ʾ�����������Զ��ύ�ļ�¼���й����ظ��ȡ���û���õļ�¼����ɾ��������</font>")
			  .Write ("</td>")
			  .Write ("</tr>")			
			  .Write ("<tr> ")
			  .Write ("<td height=""50"" colspan=""3""  align=""right"">")
			  Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.Form.asp", True, "��",CurrentPage, "action=result&itemID=" & ID)
			  .Write ("<br></td>")
			  .Write ("</tr>")			
			  .Write "</table>"
		 %>
		 <form name="export" action="?action=export" method=post target="_blank">
		  <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
                  <input type="hidden" value="<%=id%>" name="id">
		   <strong>��ʱ��ε���Excel</strong>
		   ��ʼʱ��:<input type="text" name="startdate" size="16" value="<%=dateadd("d",now,-30)%>">
		   ����ʱ��:<input type="text" name="enddate" size="16" value="<%=formatdatetime(now,2)%>">
		   <input type="submit" class="button" value="����Excel">
		   <input type="button" class="button" value="ȫ������Excel" onClick="window.open('?action=export&id=<%=id%>')">
		  </div>
		  </form>
		 
		 <%
			  .Write "</div>"
         end with
		End Sub
		
		
		'������ʾ
		Sub SubmitResultHP()
		ID=KS.ChkClng(KS.G("itemID"))
		Dim TableName,SQL,II
		TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & ID)
		MaxPerPage = 20     'ȡ��ÿҳ��ʾ����
		If KS.G("page") <> "" Then
			  CurrentPage = KS.ChkClng(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		 with response
		 %>
		  <script>
			function ShowReplay(formid,id)
			{  
			onscrolls=false;  //ȡ������
			new KesionPopup().PopupCenterIframe("�ظ�����¼","KS.Form.asp?Action=replay&formid="+formid+"&id=" +id,650,400,'no')
			}
			</script>
			<div style="margin:8px"><a href="KS.Form.asp?ItemID=<%=id%>&action=resulthp"><strong>������ʾ��¼</strong></a> | <a href="KS.Form.asp?ItemID=<%=id%>&action=result"><strong>������ʾ��¼</strong></a></div>
		
		<div style="width:800px;padding-bottom:5px;margin-bottom:5px;overflow-x: auto; height:auto">
		 <%
 			set rs=server.createobject("adodb.recordset")
			rs.open "Select Title,FieldName From KS_FormField Where ShowOnManage=1 And ItemID=" & ID & " Order By OrderID,FieldID",Conn,1,1
			If Not RS.Eof Then SQL=RS.GetRows(-1)
			RS.Close

		 	.Write "<table cellspacing=""1"" bordercolor=""#000000"" bgcolor=""#000000""  width='100%' align='center'>"
			.Write "<form name=""form1"" action=""KS.Form.asp?ItemID=" & ID &""" method=""post"">"
			.Write "<input type='hidden' name='action' id='action' value='setstatus'/>"
			.Write "<input type='hidden' name='v' id='v' value='2'/>"
			.Write "<input type='hidden' name='formid' id='formid' value='" & ID &"'/>"
			.Write "<tr height='25' bgcolor='#ffffff'>"
			.Write "  <td width='40' align='center' nowrap>ѡ��</td>"
			If IsArray(SQL) Then
				For ii=0 To Ubound(SQL,2)
				  .Write "<td align='center' nowrap>" & SQL(0,II) & "</td>"
				Next
			End If
			.Write "<td align=center nowrap>�ύʱ��</td>"
			.Write "<td align=center nowrap>״̬</td>"
			.Write "<td align=center nowrap>���������</td>"
			.Write "</tr>"
			 rs.open "select * from " & TableName & " order by adddate desc" ,conn,1,1
			 If Not RS.EOF Then
					        totalPut = Conn.Execute("Select count(1) From " & TableName)(0)
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If CurrentPage <> 1 Then
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								Else
									CurrentPage = 1
								End If
							End If
							
							  dim k,i:i=1
							   dim rsf:set rsf=server.CreateObject("adodb.recordset")
							  do while not rs.eof
							   response.write "<tr bgcolor='#ffffff'><td width=40 align='center' nowrap><input type='checkbox' name='id' value='" & rs("id") & "'></td>"
							   If IsArray(SQL) Then
							    For II=0 To Ubound(SQL,2)
								  response.write "<td>&nbsp;" & rs(trim(sql(1,ii))) & "</td>"
								Next
							   End If
								.Write "<td align=center nowrap>" & formatdatetime(rs("adddate"),2) & "</td>"
								.Write "<td align=center nowrap>"
								select case rs("status")
							   case 0
								response.write "<font color=red>δ��</font>"
							   case 1
								response.write "<font color=green>�Ѷ�</font>"
							   case 2
								response.write "<font color=#ff6600>����</font>"
							   case 3
								response.write "����"
							   end select
								.Write "</td>"
							   response.write "<td  class='splittd' nowrap>"
							   
							   response.write "<a href=""?action=delinfo&FormID=" & ID&"&id=" & rs("Id") & """ onclick=""return(confirm('ȷ��ɾ����?'))"">ɾ</a> <a href=""?action=modifyinfo&FormID=" & ID&"&id=" & rs("Id") & """>��</a> <a href='?action=setstatus&v=1&FormID=" & ID&"&id=" & rs("id") & "' title='��Ϊ�Ѷ�'>�Ѷ�</a> <a href='?action=setstatus&v=2&FormID=" & ID&"&id=" & rs("id") & "' title='��Ϊ����'>����</a> <a href='?action=setstatus&v=3&FormID=" & ID&"&id=" & rs("id") & "' title='��Ϊ����'>����</a> <a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");"">�ظ�</a>"
							   if not isnull(rs("note")) and rs("note")<>"" then response.write "&nbsp;&nbsp;<a href=""javascript:ShowReplay(" & ID& "," & rs("id") & ");""><font color=blue>�ѻظ�</font></a>"

							   response.write "</td>"
							   response.write "</tr>" 
							  rs.movenext
							  i=i+1
							  if i>maxperpage then exit do
							  loop

			 End If
			  .Write "<tr><td height='36' colspan=100 bgcolor='#ffffff'><label><input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"">ѡ��</label> <input type='submit' class='button' value='��������' onclick=""$('#action').val('setstatus');$('#v').val(2);""/> <input type='submit' class='button' value='�������ó���Ч��¼' onclick=""$('#action').val('setstatus');$('#v').val(3);""/>  <input type='submit' class='button' value='�������ó��Ѷ�' onclick=""$('#action').val('setstatus');$('#v').val(1);""/> <input type='submit' class='button' value='����ɾ��' onclick=""if (confirm('�˲���������,ȷ��ɾ��ѡ�еļ�¼��?')){$('#action').val('delinfo');}else{return false}""/> <input type='button' class='button' onclick='window.print();' value='��ӡ��ҳ���¼'></td></tr></form>"
			  .Write ("</table>")
			  .Write ("<br/>")
			   Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			  .Write ("<br>")
		 %>
		<SCRIPT language=javascript>
		function unselectall()
		{
			if(document.myform.chkAll.checked){
			document.myform.chkAll.checked = document.myform.chkAll.checked&0;
			} 	
		}
		
		function CheckAll(form)
		{
		  for (var i=0;i<form.elements.length;i++)
			{
			var e = form.elements[i];
			if (e.Name != "chkAll"  && e.disabled==false)
			   e.checked = form.chkAll.checked;
			}
		}
		</SCRIPT>
		 		 <div style="clear:both"></div>
        </div>
		 
		 <form name="export" action="?action=export" method=post target="_blank">
		  <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
                  <input type="hidden" value="<%=id%>" name="id">
		   <strong>��ʱ��ε���Excel</strong>
		   ��ʼʱ��:<input type="text" name="startdate" size="16" value="<%=dateadd("d",now,-30)%>">
		   ����ʱ��:<input type="text" name="enddate" size="16" value="<%=formatdatetime(now,2)%>">
		   <input type="submit" class="button" value="����Excel">
		   <input type="button" class="button" value="ȫ������Excel" onClick="window.open('?action=export&id=<%=id%>')">
		  </div>
		  </form>
		 
		 <%
			  .Write "</div>"
         end with
		End Sub
		
		'�޸ļ�¼
		Sub modifyinfo()
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		  Dim FormID:FormID=KS.ChkClng(KS.S("FormID"))
		  Dim Title,TableName,SQL,ii
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 FormName,TableName From KS_Form Where ID=" & FormID,Conn,1,1
		  If RS.Eof And RS.Bof Then
		    RS.Close :Set RS=Nothing
			KS.AlertHintScript "�Բ���,������!"
		  End If
		  Title=RS(0) : TableName=RS(1)
		  RS.Close 
		  RS.Open "Select Title,FieldName,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,AllowFileExt,MaxFileSize,FieldID From KS_FormField Where ItemID="& FormID,conn,1,1
		  If Not RS.Eof Then SQL=RS.GetRows(-1)
		  RS.Close
		  
		  If ID<>0 Then
			  RS.Open "Select top 1 * From " & TableName & " Where ID=" & ID,conn,1,1
			  If RS.Eof And RS.Bof Then
				RS.Close :Set RS=Nothing
			  End If
		  End If
		  %>
		  <div style="margin:6px;text-align:center">
		  <%if id=0 then
		    response.write "���"
			else
			response.write "�޸�"
			end if
		%>��[<span style='color:red'><%=Title%></span>]���ύ��¼</div>
		  <table width='99%' align="center" class="ctable" border='0' cellspacing='1' cellpadding='1'>
		   <form name="myform" action="KS.Form.asp" method="post">
		   <input type="hidden" value="DoResultSave" name="action">
		   <input type="hidden" value="<%=ID%>" name="id">
		   <input type="hidden" value="<%=formid%>" name="formid">
		    <%
			If IsArray(SQL) Then
			   For II=0 To Ubound(SQL,2)
			 %>
		  <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong><%=SQL(0,II)%>��</strong></div></td>      
			<td height="30"> 
			
			<%
			Dim O_Arr,O_Len,n,F_V,O_Value,O_Text,BRStr,FieldValue
			if ID<>0 Then
			FieldValue=RS(Trim(SQL(1,II)))
			Else
			FieldValue=SQL(4,II)
			End If
			Select Case SQL(3,ii)
				Case 2
				  Response.Write "<textarea style=""width:" & SQL(7,ii) & "px;height:" & SQL(8,ii) &"px"" rows=""5"" name=""" & SQL(1,ii) & """>" & FieldValue & "</textarea>"
			   Case 3
				  Response.Write "<select class=""upfile"" style=""width:" & SQL(7,ii) & "px"" name=""" & SQL(1,ii) & """>"
				  O_Arr=Split(SQL(5,ii),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
					If O_Arr(N)<>"" Then
						F_V=Split(O_Arr(N),"|")
						If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						Else
							O_Value=F_V(0):O_Text=F_V(0)
						End If						   
						If trim(FieldValue)=trim(O_Value) Then
							Response.Write "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
						Else
							Response.Write "<option value=""" & O_Value& """>" & O_Text & "</option>"
						End If
					End If
				  Next
				  Response.Write "</select>"
			  Case 6
				  O_Arr=Split(SQL(5,ii),vbcrlf): O_Len=Ubound(O_Arr)
				  If O_Len>1 And Len(SQL(5,I))>50 Then BrStr="<br>" Else BrStr=""
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If trim(FieldValue)=trim(O_Value) Then
						Response.Write "<input type=""radio"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
					Else
						Response.Write "<input type=""radio"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """>" & O_Text & BRStr
					End If
				   End If
				  Next
			 Case 7
				   O_Arr=Split(SQL(5,ii),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
				    If O_Arr(N)<>"" Then
					F_V=Split(O_Arr(N),"|")
					If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					Else
						O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If KS.FoundInArr(trim(FieldValue),O_Value,",")=true Then
						Response.Write "<input type=""checkbox"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """ checked>" & O_Text
					Else
						Response.Write "<input type=""checkbox"" name=""" & SQL(1,ii) & """ value=""" & O_Value& """>" & O_Text
					End If
					End If
				  Next
			 Case 10
					Response.Write "<textarea id=""" & SQL(1,ii) &""" name=""" & SQL(1,ii) &""">"& Server.HTMLEncode(trim(FieldValue)) &"</textarea>"
					%>
				 <script type="text/javascript" src="../editor/ckeditor.js"></script>
		   <script type="text/javascript">
                CKEDITOR.replace('<%=SQL(1,ii)%>', {width:"100%",height:"150px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			</script>
			<%
			 Case Else
				Response.Write "<input type=""text"" class=""upfile"" style=""width:" & SQL(7,ii) & "px"" name=""" & SQL(1,ii) & """ value=""" & FieldValue & """>"
			End Select
			%>
			
			</td> 
		 </tr>
		    <%Next
		   End If
		   %>
		   <tr> 
		    <td class='tdbg' colspan=3 style="text-align:center">
			  <input type="hidden" name="comeurl" value="<%=Request.ServerVariables("HTTP_REFERER")%>"/>
			  <input type="submit" value="�ύ����" class="button"/>
			</td>
		   </tr>
		   </form>
		  </table>
		  <br/><br/>
		  <%
		  if ID<>0 Then
			  RS.Close
			  Set RS=Nothing
		  End If
		End Sub
		
		'������ύ������޸�
		Sub DoResultSave()
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		  Dim FormID:FormID=KS.ChkClng(KS.S("FormID"))
		  Dim Title,TableName,SQL,ii
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 FormName,TableName From KS_Form Where ID=" & FormID,Conn,1,1
		  If RS.Eof And RS.Bof Then
		    RS.Close :Set RS=Nothing
			KS.AlertHintScript "�Բ���,������!"
		  End If
		  Title=RS(0) : TableName=RS(1)
		  RS.Close 
		  RS.Open "Select Title,FieldName From KS_FormField Where ItemID="& FormID,conn,1,1
		  If Not RS.Eof Then SQL=RS.GetRows(-1)
		  RS.Close
		  RS.Open "Select top 1 * From " & TableName & " Where ID=" & ID,conn,1,3
		  If RS.Eof And RS.Bof Then
		    RS.AddNew
			RS("Status")=1
			RS("AddDate")=Now
		  End If
		  For Ii=0 To Ubound(SQL,2)
		    RS(Trim(SQL(1,II)))=KS.G(Trim(SQL(1,II)))
		  Next
		   RS.Update
		   RS.Close
		   Set RS=Nothing
		   if id=0 then
			   Response.Write "<script>alert('��ϲ,��ӳɹ�!');location.href='KS.Form.asp?ItemID=" & FormID&"&action=resulthp';</script>"
		   else
			   If KS.G("ComeUrl")<>"" Then
			   Response.Write "<script>alert('��ϲ,�޸ĳɹ�!');location.href='" & Request("comeurl") &"';</script>"
			   Else
			   Response.Write "<script>alert('��ϲ,�޸ĳɹ�!');location.href='KS.Form.asp?ItemID=" & FormID&"&action=resulthp';</script>"
			   End If
		   end if
		 
		End Sub
		
		Sub Replay()
		 on error resume next
		 Dim FormID:FormID=KS.ChkClng(KS.G("FormID"))
		 Dim ID:ID=KS.ChkClng(KS.G("id"))
		 Dim TableName:TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & FormID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From " & TableName &" Where ID=" & ID,conn,1,1
		 If RS.Eof Then
		  response.end
		 End If
         %>
		 <iframe src="about:blank" style="display:none" name="hiddenframe"></iframe>
		 <form action="KS.Form.asp?action=replaysave&formid=<%=formid%>&id=<%=id%>" method="post" name="myform" target="hiddenframe">
		  <br>
		  <div style="margin:6px;text-align:center;font-weight:bold;color:red">�鿴�ظ�</div>
		  <table width='99%' align='center' border='0' cellpadding='1'  cellspacing='1' class='ctable'> 
		  <tr class="tdbg">
		    <td align="right" class="clefttitle">����ʱ��</td>
			<td><%=rs("adddate")%></td>
		  </tr>
		    
		  <tr class="tdbg">
		   <td align="right" class="clefttitle">�ظ����ݣ�</td>
		   <td>
			<textarea id="content" name="content"><%=server.htmlencode(rs("note"))%></textarea>
			 <script type="text/javascript" src="../editor/ckeditor.js"></script>
		   <script type="text/javascript">
                CKEDITOR.replace('content', {width:"100%",height:"150px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			</script>	

		   </td>
		  </tr>
		  <tr class="tdbg">
		    <td align="right" class="clefttitle">�����ʼ�</td>
			<td><label><input type="checkbox" name="sendmail" value="1" checked="checked">���ظ����ݷ��͵��û�����</label>
			
			&nbsp;<span style='color:#999999'>���ʱҪ��ͻ����ʼ��Ĳ���Ч��</span>
			</td>
		  </tr>
		  <tr  class="tdbg">
		    <td colspan="2" height="35" style="text-align:center"><input type="submit" class="button" value="�ύ�ظ�">&nbsp;<input type="button" class="button" value="�رմ���" onClick="parent.closeWindow()"></td>
		  </tr>
		  </table>
		 </form>
		 <%
		  RS.Close:Set RS=Nothing
		End Sub
		
		Sub setstatus()
		 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
		 If Id="" Then KS.AlertHintScript "�Բ���,��û��ѡ��!"
		 conn.execute("update " & LFCls.GetSingleFieldValue("Select TableName From KS_Form Where ID=" & KS.ChkClng(KS.G("FormID"))) &" set status=" & ks.chkclng(ks.g("v")) & " where id in(" & id &")")
		 response.redirect request.servervariables("http_referer")
		End Sub
		
		Sub DelInfo()
		 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
		 If Id="" Then KS.AlertHintScript "�Բ���,��û��ѡ��!"
		 conn.execute("delete from " & LFCls.GetSingleFieldValue("Select TableName From KS_Form Where ID=" & KS.ChkClng(KS.G("FormID"))) &" where id in (" & id &")")
		 response.redirect request.servervariables("http_referer")
		End Sub
		
		Sub TemplateSave()
		 Dim FormID,TContent,K
		 FormID=KS.ChkCLng(KS.G("FormID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select StepNum,PostByStep,Template From KS_Form Where ID=" & FormID,conn,1,3
		 If Not RS.Eof Then
		   If RS(1)=1 Then
	   		 For K=1 To RS("StepNum")
			  If K=1 Then
			  Tcontent=Request.Form("Content"&K)
			  Else
			  Tcontent=Tcontent & "$aaa$" & Request.Form("Content"&K)
			  End If
    		 Next
		   Else
		     Tcontent=Request.Form("Content1")
		   End IF
		   RS(2)=Tcontent
		  RS.Update
		 End If
		 RS.Close:Set RS=Nothing
		 Response.Write"<script>alert('��ϲ��ģ���޸ĳɹ�!');location.href='KS.Form.asp';</script>"
		End Sub
		
		Sub ReplaySave()
		 Dim FormID:FormID=KS.ChkClng(KS.G("FormID"))
		 Dim ID:ID=KS.ChkClng(KS.G("id"))
		 Dim TableName:TableName=LFCls.GetSingleFieldValue("Select top 1 TableName From KS_Form Where ID=" & FormID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select note From " & TableName &" Where ID=" & ID,conn,1,3
          RS(0)=Request.Form("Content")
		 RS.Update
		 RS.Close
		 
		 If KS.ChkClng(KS.G("SendMail"))=1 Then   '���ʼ�֪ͨ
		  Dim EmailField,Email,FormName
		  Set RS=Conn.Execute("Select top 1 FieldName From KS_FormField Where FieldType=8 and ItemID=" & FormID)
		  If Not RS.Eof Then
		     EmailField=RS(0)
			 RS.Close
			 Set RS=Conn.Execute("Select Top 1 " & EmailField & " From " & TableName & " Where ID=" & ID)
			 If Not RS.Eof Then
			    Email=RS(0)
			 End If
			 If  KS.IsValidEmail(Email) Then
			    RS.Close
			   Dim S_Content,sql,k,ReturnInfo,UpFiles
			   set rs=conn.execute("select FieldName,title,MustFillTF,FieldType,ShowUnit from ks_formfield where itemid=" & Formid & " and ShowOnForm=1 order by orderid")
			   sql=rs.getrows(-1)
			   rs.close
			   rs.open "select top 1 * From " & TableName & " Where ID=" & ID,conn,1,1
			   s_content="<table border=0 cellpadding=0 cellspacing=0>" & vbcrlf
			   for k=0 to ubound(sql,2)
				
				s_content=s_content &"<tr>" & vbcrlf
				s_content=s_content & "<td width=120 align=right>" & sql(1,k) & ":</td>" & vbcrlf
				s_content=s_content & "<td>" 
				
				s_content=s_content & rs(trim(sql(0,k)))
				
				s_content=s_content & "</td>" & vbcrlf
				s_content=s_content & "</tr>" & vbcrlf
			   next
				s_content=s_content &"</table>"
				
				FormName=Conn.Execute("select top 1 formname from ks_form where id=" & formid)(0)
				s_content="�𾴵��û������ã�<br />&nbsp;&nbsp;&nbsp;&nbsp;����������<font color=""red"">"  &KS.Setting(0) & "</font>�ύ[" & FormName & "]����Ϣ:<br />" & s_content & "<br/><strong>�����Ǳ�վ����Ա�����Ĵ𸴣�</strong>" & Request.Form("Content")
				
				ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14),KS.Setting(0) & "�������ύ[" & FormName & "]�Ļظ�!", Email,KS.Setting(0), s_content,KS.Setting(11))
			   If ReturnInfo="OK" Then
				ReturnInfo="�ѽ��ύ������͵���������" & Email & "!"
			   Else
				ReturnInfo=""
			   End If
		   
			 
			 End If
		  End If
		  RS.Close
		 End If
		 Set RS=Nothing
		 If ReturnInfo<>"" Then
		 Response.Write "<script>alert('��ϲ���ύ�ظ��ɹ���" & ReturnInfo & "');parent.parent.location.reload();</script>"
		 Else
		 Response.Write "<script>alert('��ϲ���ύ�ظ��ɹ���');parent.parent.location.reload();</script>"
		 End If
		End Sub
		
		Sub export()
		    dim param
			Dim id:id=ks.chkclng(request("id"))
			dim startdate:startdate=request("startdate")
			dim enddate:enddate=request("enddate")
			if id=0 then ks.die "error!"
			
			Dim TableName:TableName=LFCls.GetSingleFieldValue("Select TableName From KS_Form Where ID=" & ID)
			
			param=" where 1=1"
			
			if startdate<>"" and not isdate(startdate) then
				 response.write "<script>alert('��ʼʱ���ʽ����ȷ!');window.close();</script>"
				 response.end
			end if
			if enddate<>"" and not isdate(enddate) then
				 response.write "<script>alert('����ʱ���ʽ����ȷ!');window.close();</script>"
				 response.end
			end if
			
				if isdate(startdate) and isdate(enddate) then
				 EndDate = DateAdd("d", 1, enddate)
				 if DataBaseType=1 then
				 param=param &" and AddDate>= '" & StartDate & "' And  AddDate <='" & EndDate & "'"
				 else
				 param=param &" and AddDate>= #" & StartDate & "# And  AddDate <=#" & EndDate & "#"
				 end if
				else
				end if
			
			
			Response.AddHeader "Content-Disposition", "attachment;filename=addressbook.xls" 
			Response.ContentType = "application/vnd.ms-excel" 
			Response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=GB2312"">"
			
			dim sql,i
			
			dim rs:set rs=server.CreateObject("adodb.recordset")
			rs.open "Select title,fieldname From [KS_FormField] Where ItemID=" & ID & " Order by OrderID",conn,1,1
			if not rs.eof then
			 sql=rs.getrows(-1)
			end if
			rs.close
			if not isarray(sql) then
			 response.write "<script>alert('û�м�¼!');window.close();</script>"
			end if
			
			response.write "<table width=""100%"" border=""1"" >" 
			response.write "<tr>" 
			for i=0 to ubound(sql,2)
			response.write "<th><b>" & sql(0,i) & "</b></th>" 
			next
			response.write "<th><b>�û���</b></th>"
			response.write "<th><b>�ύʱ��</b></th>"
			response.write "</tr>" 
			
			rs.open "select  * from " & TableName & " " & param & " order by id desc",conn,1,1
			do while not rs.eof
			  
			  response.write "<tr>"
			  for i=0 to ubound(sql,2) 
			  response.write "<td align=center>" & ks.htmlcode(rs(sql(1,i))) & "&nbsp;</td>" 
			  next 
			  response.write "<td align=center>" & rs("username") & "</td>"
			  response.write "<td align=center>" & rs("adddate") & "</td>"
			  response.write "</tr>" 
			  rs.movenext
			loop
			rs.close
			
			
			response.write "</table>"

		End Sub
		
		Sub Import()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim Title
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Form Where ID=" & id,Conn,1,1
		 If RS.Eof And RS.Bof Then 
		   RS.Close:Set RS=Nothing
		   KS.AlertHintScript "����������!"
		 End If
		 Title=RS("FormName")
		 RS.Close :Set RS=Nothing
		%>
		<div class="sort" style="line-height:30px">��������Excel���ݵ���[<font color=red><%=title%></font>]</div>
			<table width="100%" style="margin-top:10px" border="0" align="center"  cellspacing="1" class='ctable'>
			<form name="myform" action="?Action=ImportNext" method="post">
			<input type="hidden" name="id" value="<%=id%>"/>
			<input type="hidden" name="title" value="<%=title%>"/>
			  
			  <tr class='tdbg'> 
			    <td height="25" align='right' class='clefttitle'><strong>ѡ��Ҫ�����Excel�ļ�:</strong></td>
				<td><input name='FilePath' type='text' class='upfile' id='FilePath' value="/1.xls" size=20></td>
              </tr>
			  <tr class='tdbg'> 
			    <td height="25" align='right' class='clefttitle'><strong>����Excel�ı�����:</strong></td>
				<td><input name='tablename' type='text' class='upfile' id='tablename' value="Sheet1" size=20></td>
              </tr>
		 <tr class='tdbg'>
		    <td colspan=2 height='30'><b>˵����</b>
			<br/>
			1���뽫Ҫ�����Excel�ļ��ϴ�����վ�ϣ�Ȼ��������ȷ��Excel·����
			<br/>
			2���밴��ʽ�����excel���ݣ���ʽ���£�
			<br/>
			<!--<div style="width:800px;padding-bottom:35px;overflow-x: auto; height:auto">-->
			<table width="100%" border="1" ><tr>
			<%
			Dim SQL,ii
			 Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select Title,FieldName From KS_FormField Where ItemID=" & ID,Conn,1,1
			 If Not RS.Eof Then SQL=RS.GetRows(-1)
			 RS.Close : Set RS=Nothing
			 If IsArray(SQL) Then
			   For II=0 To Ubound(SQL,2)
			    response.write "<th><b>" & sql(0,ii) &"</b></th>"
               Next
			 End If
			%>
			</tr></table>
			<!--</div>-->
			<br/>
			
			<br><div align='center'> <input type="submit" class="button" name="button1" value="��һ��"> 
				  &nbsp; <input type="reset" class="button" name="button2" value=" ���� "> </div></td>
		 </tr>
			  </form>
			</table>
		<%
		End Sub
		
Sub OpenImporIConn()
		   if not isobject(IConn) then
			on error resume next
		    Set IConn = Server.CreateObject("ADODB.Connection")
			IConn.open IConnStr
			If Err Then 
			  Err.Clear
			  Set IConn = Nothing
			  Response.Write "<script>alert('Excel���ݿ�����ʧ��!');history.back();</script>"
			  response.end
			end if
		   end if		
End Sub
'**************************************************
	'��������ShowChird
	'��  �ã���ʾָ�����ݱ���ֶ��б�
	'��  ������
	'**************************************************
	Sub ShowField(fieldname)
	        if request("tablename")="" then
			 response.write "<script>alert('�����Ʊ������룡');history.back();</script>"
			 response.end
			end if
	        dim dbname:dbname=request("tablename") & "$"
		    dim rs:Set rs=Iconn.OpenSchema(4)
			Do Until rs.EOF or rs("Table_name") = trim(dbname)
				rs.MoveNext
			Loop
            Do Until rs.EOF or rs("Table_name") <> trim(dbname)
			  if fieldname=trim(rs("column_name")) then
				response.write "<option value='"&rs("column_Name")&"' selected>��"&rs("column_Name")&"</option>"
			  else
				response.write "<option value='"&rs("column_Name")&"'>��"&rs("column_Name")&"</option>"
			  end if
					   rs.MoveNext
			loop
			rs.close:set rs=nothing
	End Sub

Sub importNext()
         Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 If ID=0 Then KS.Die "<script>alert('error!');history.back();</script>"
         Dim FilePath:FilePath=Request.Form("FilePath")
		 IConnStr="driver={microsoft excel driver (*.xls)};ReadOnly=1;dbq=" & Server.Mappath(FilePath)
		 OpenImporIConn()
		 %>
		 	<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; chaRSet=gb2312">
			<link href="Include/Admin_Style.css" rel="stylesheet">
			<script language="JavaScript" src="Include/Common.js"></script>
           </head>
			<body topmargin="0" leftmargin="0">
			<div class="sort" style="line-height:30px">�����������ݵ���[<font color=red><%=KS.S("Title")%></font>](���õ�����)</div>
			<table width="100%" style="margin-top:10px" border="0" align="center"  cellspacing="1" class='ctable'>
			<form name="myform" action="?Action=ImportNext2" method="post">
			 <input type="hidden" value="<%=id%>" name="id">
			 <input type="hidden" value="<%=FilePath%>" name="FilePath">
			 <input type="hidden" value="<%=request("tablename")%>" name="tablename">
			 <input type="hidden" value="<%=request("title")%>" name="title">
			 <%
			 Dim RS,SQL,ii
			 Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select Title,FieldName From KS_FormField Where ItemID=" & ID,Conn,1,1
			 If Not RS.Eof Then SQL=RS.GetRows(-1)
			 RS.Close : Set RS=Nothing
			 If IsArray(SQL) Then
			   For II=0 To Ubound(SQL,2)
			 %>
			  <tr class='tdbg'> 
			    <td height="25" align='right' class='clefttitle'>
				<select name='<%=SQL(1,II)%>_Y'>
				<option value="0">-�������-</option>
				<%ShowField(SQL(0,II))%>
				</select> =>	</td>
				<td><%=SQL(0,II)%>(<%=SQL(1,II)%>)</td>
			  </tr>
			 <%Next
			 End If
			 %> 
			  
		 <tr class='tdbg'>
		    <td colspan=2 height='30'><br/><b>˵����</b>����ȷ���������ֶζ�Ӧ,Ȼ�����һ����ʼ���������<br/><br><div align='center'> <input type="submit" class="button" name="button1" value="��һ��"> 
				  &nbsp; <input type="reset" class="button" name="button2" value=" ���� "> </div></td>
		 </tr>
			  </form>
			</table>
			</body>
			</html>
<%
end sub

Sub ImportNext2()
%>
<div class="sort" style="line-height:30px">�����������ݵ���[<font color=red><%=Request("Title")%></font>](����ִ�е���)</div>
		<div style="text-align:center">			 
			 <div style="margin-top:50px;border:1px dashed #cccccc;width:500px;height:80px">
			 <br>
			<div id="message">
			  <br>������ʾ����
			</div>
			</div>
	    </div>
		<br/><br/><br/>
	 <%
	     'On Error Resume Next
		 Server.ScriptTimeOut=999999
	     Dim TableName:TableName="[" & request("tablename") & "$]"
		 Dim N,FoundErr,Total,ErrNum:ErrNum=0
		 Dim t:t=0
		 Dim SQL,II,msg,ToTableName
         Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 If ID=0 Then KS.Die "<script>alert('error!');history.back();</script>"
		 Dim FilePath:FilePath=Request.Form("FilePath")
		 IConnStr="driver={microsoft excel driver (*.xls)};ReadOnly=1;dbq=" & Server.Mappath(FilePath)
		 OpenImporIConn()
		 Dim IRS:Set IRS=Server.CreateOBject("ADODB.RECORDSET")
    	 Dim RS:Set RS=Server.CreateObject("ADODB.RecordSet")
		 RS.Open "Select Top 1 TableName From KS_Form Where ID=" & ID,Conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close :Set RS=Nothing
		   KS.AlertHintScript "error!"
		 End If
		 ToTableName=RS(0)
		 RS.Close
		 RS.Open "Select Title,FieldName From KS_FormField Where ItemID=" & ID,Conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close 
			 
		 IRS.Open "Select * From " & TableName,iConn,3,3

		 Total=IRS.RecordCount
		 Do While Not IRS.Eof
		   
		   t=t+1
		   FoundErr=false
		   
		'   If founderr=true Then
		'   	  response.write msg 
		 '  end if
		   
				 RS.Open "Select top 1 * From " & ToTableName &" Where 1=0",conn,1,3
				 If RS.Eof and RS.Bof Then
				   RS.AddNew
				   
				   If IsArray(SQL) Then
				     For II=0 To Ubound(SQL,2)
					  if Request(trim(SQL(1,II)) & "_y")<>"0" then
				       'response.write "RS(" & trim(SQL(1,II) & ")=IRS(" & Request(trim(SQL(1,II)) & "_y")) & ")<br/>"
				       RS(trim(SQL(1,II)))=IRS(trim(Request(trim(SQL(1,II)) & "_y")))
					  end if
					 Next
				   End If
                   RS("AddDate")=Now
				   RS("status")=1
				   RS.Update
					 N=N+1
				Else
				 ErrNum=ErrNum+1
				End If
				RS.Close
		    'Else
			'   ErrNum=ErrNum+1
			'End If
		  	Response.Write "<script>document.all.message.innerHTML='<br>��<font color=red>" & Total & "</font> �����ݣ����ڵ����<font color=red>" & n & "</font>������������<font color=blue>" & ErrNum & "</font>��!';</script>"
			Response.Flush

		  IRS.MoveNext
		  If t>=Total Then Exit Do
		 Loop
		 IRS.Close:Set IRS=Nothing:Set RS=Nothing
		 Response.Write "<script>document.all.message.innerHTML='<br>��ϲ���ɹ����� <font color=red>" & N & "</font> �����ݣ�����" & errnum &" ��';</script>"
		 
		 if msg<>"" then
		   response.write "<strong>���¼�¼�ظ�û���ٵ���:</strong><br/><font color=red>" & msg & "</font>"
		 end if
		 
End Sub	
		
		
End Class
%> 
