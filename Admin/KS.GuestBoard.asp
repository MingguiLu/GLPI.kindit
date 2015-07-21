<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.FunctionCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GuestBoard_Main
KSCls.Kesion()
Set KSCls = Nothing

Class GuestBoard_Main
        Private KS,Action
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 10
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMS20004") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Action=KS.G("Action")
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "var Page='" & CurrentPage & "';" & vbCrLf
			.Write "</script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"

			Select Case Action
			 Case "Add","Edit" Call GuestBoardAddOrEdit()
			 Case "Save" Call GuestBoardSave()
			 Case "Del" Call GuestBoardDel()
			 Case "DelTopic" Call DelTopic()
			 Case Else
			   Call MainList()
			End Select
		  End With
	    End Sub
		
		Sub MainList()
		 With Response
			%>
			<script language="JavaScript">
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			function document.onreadystatechange()
			{   if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','GuestBoardID');
				InitialContextMenu();
				DocElementArrInitialFlag=true;
			}
			function InitialContextMenu()
			{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("window.parent.GuestBoardAdd(0);",'�� ��(N)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.GuestBoardControl(1);",'�� ��(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.GuestBoardControl(2);",'ɾ ��(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
			}
			function DocDisabledContextMenu()
			{
				DisabledContextMenu('FolderID','GuestBoardID','�� ��(E),ɾ ��(D)','�� ��(E)','','','','')
			}
			function GuestBoardAdd(parentid)
			{
				location.href='KS.GuestBoard.asp?Action=Add&parentid='+parentid;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=����������� >> <font color=red>����°���</font>&ButtonSymbol=GO';
			}
			function EditGuestBoard(id)
			{
				location="KS.GuestBoard.asp?Action=Edit&Page="+Page+"&Flag=Edit&GuestBoardID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=����������� >> <font color=red>�༭����</font>&ButtonSymbol=GoSave';
			}
			function DelGuestBoard(id)
			{
			if (confirm('������Ӱ��潫ͬʱ��ɾ��,���Ҫִ��ɾ��������?'))
			 location="KS.GuestBoard.asp?Action=Del&Page="+Page+"&GuestBoardid="+id;
			   SelectedFile='';
			}
			function DelTopic(id){
			if (confirm('ִ�д˲�������ոð汾�����������ͻظ�,�˲��������������ز���!!!'))
			 location="KS.GuestBoard.asp?Action=DelTopic&Page="+Page+"&GuestBoardid="+id;
			   SelectedFile='';
			}
			function GuestBoardControl(op)
			{  var alertmsg='';
				GetSelectStatus('FolderID','GuestBoardID');
				if (SelectedFile!='')
				 {  if (op==1)
					{
					if (SelectedFile.indexOf(',')==-1) 
						EditGuestBoard(SelectedFile)
					  else alert('һ��ֻ�ܱ༭һ������!')	
					SelectedFile='';
					}
				  else if (op==2)    
					 DelGuestBoard(SelectedFile);
				 }
				else 
				 {
				 if (op==1)
				  alertmsg="�༭";
				 else if(op==2)
				  alertmsg="ɾ��"; 
				 else
				  {
				  WindowReload();
				  alertmsg="����" 
				  }
				 alert('��ѡ��Ҫ'+alertmsg+'�İ���');
				  }
			}
			function GetKeyDown()
			{ 
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 : location.reload(); break;
				 case 65 : SelectAllElement();break;
				 case 78 : event.keyCode=0;event.returnValue=false; GuestBoardAdd(0);break;
				 case 69 : event.keyCode=0;event.returnValue=false;GuestBoardControl(1);break;
				 case 68 : GuestBoardControl(2);break;
			   }	
			else	
			 if (event.keyCode==46)GuestBoardControl(2);
			}
			</script>
			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""GuestBoardAdd(0);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��Ӱ���</span></li>"
			  .Write "<li class='parent' onclick=""GuestBoardControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭����</span></li>"
			  .Write "<li class='parent' onclick=""GuestBoardControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ������</span></li>"
			  .Write "</ul>"
			

			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr>"			
			.Write "          <td height=""25"" class=""sort"" align=""center"">��������</td>"
			.Write "          <td class=""sort""><div align=""center"">����</div></td>"
			.Write "          <td align=""center"" class=""sort"">������</td>"
			.Write "          <td width=""50"" class=""sort"" align=""center"">����</td>"
			.Write "          <td class=""sort"" align=""center"">�������</td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
			 SqlStr = "SELECT * FROM KS_GuestBoard Where ParentID=0 order by orderID,id"
			 RSObj.Open SqlStr, Conn, 1, 1
			 If RSObj.EOF And RSObj.BOF Then
			 Else
						        totalPut = RSObj.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
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
			  Dim RS,I
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr>"
					  .Write "  <td class='splittd' height='20'>&nbsp; <span GuestBoardID='" & RSObj("ID") & "' ondblclick=""EditGuestBoard(this.GuestBoardID)""><img src='Images/Field.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & RSObj("BoardName") & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' align='center'>&nbsp;" & RSObj("master") & "&nbsp;</td>"
					  .Write "  <td class='splittd' align='center'>����:<font Color=red>" & RSObj("topicnum") & "</font> ����:<font Color=red>" & RSObj("postnum") & "</font></td>"
					  .Write "  <td class='splittd' align='center'>" & RSOBJ("OrderID") &"</td>"
					  .Write "  <td class='splittd'> <a href='javascript:GuestBoardAdd(" & rsobj("id") & ")'>��ӷְ�</a> | <a href='javascript:EditGuestBoard(" & rsobj("id") & ")'>�޸�</a> | <a href='javascript:DelGuestBoard(" & rsobj("id") & ")'>ɾ��</a> </td>"
					  .Write "</tr>"
					  Set RS=Conn.Execute("Select ID,BoardName,master,todaynum,postnum,topicnum,orderid From KS_GuestBoard Where ParentID=" & RSObj("ID") & " Order by orderid")
					  Do While not rs.eof
					  .Write "<tr>"
					  .Write "  <td class='splittd' height='20'> &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;|- <span GuestBoardID='" & RS("ID") & "' ondblclick=""EditGuestBoard(this.GuestBoardID)""><img src='Images/folder/folderopen.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & RS("BoardName") & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' align='center'>&nbsp;" & RS("master") & "&nbsp;</td>"
					  .Write "  <td class='splittd' align='center'>����:<font Color=red>" & RS("topicnum") & "</font> ����:<font Color=red>" & RS("postnum") & "</font></td>"
					  .Write "  <td class='splittd' align='center'>" & RS("OrderID") &"</td>"
					  .Write "  <td class='splittd'> <a href='#' disabled>��ӷְ�</a> | <a href='javascript:EditGuestBoard(" & rs("id") & ")'>�޸�</a> | <a href='javascript:DelGuestBoard(" & rs("id") & ")'>ɾ��</a>  | <a href='javascript:DelTopic(" & rs("id") & ")'>���</a> </td>"
					  .Write "</tr>"
					  rs.movenext
					  loop
					  rs.close
					  
					 I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  .Write "<tr><td height='26' colspan='5' align='right'>"
					  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "��", CurrentPage, "Action=" & Action)
				End With
			    Set RS=Nothing
			End Sub
			
			'����޸İ���
		  Sub GuestBoardAddOrEdit()
		  		Dim GuestBoardID, RSObj, SqlStr, Content, BoardName, Note, Master, AddDate,Flag, Page,OrderID,ParentID,BoardRules,Settings,SetArr,Locked
				Flag = KS.G("Flag")
				Page = KS.G("Page")
				If Page = "" Then Page = 1
				If Flag = "Edit" Then
					GuestBoardID = KS.G("GuestBoardID")
					Set RSObj = Server.CreateObject("Adodb.Recordset")
					SqlStr = "SELECT top 1 * FROM KS_GuestBoard Where ID=" & GuestBoardID
					RSObj.Open SqlStr, Conn, 1, 1
					  BoardName     = RSObj("BoardName")
					  Note    = RSObj("Note")
					  AddDate  = RSObj("AddDate")
					  Master  = RSObj("Master")
					  ParentID= RSObj("ParentID")
					  OrderID = RSObj("OrderID")
					  BoardRules=RSObj("BoardRules")
					  Locked = RSObj("Locked")
					  Settings=RSObj("Settings")&"$0$0$0$0$1$1$1$1$20$$1$1$10$1$0$0$0$1$1$20$20$0$0$0$0$1$1$1$1$20$$1$1$10$1$0$0$0$1$1$20$20$$$$$$$$$$$$$$$$$$$$$$$$$$"
					RSObj.Close:Set RSObj = Nothing
				Else
				   Flag = "Add"
				   ParentID=Request("Parentid")
				   BoardRules="���ް��" : Locked=0
				End If
				Settings=Settings&"0$0$0$1$1$1$1$1$1$20$$0$0$10$1$0$0$0$1$1$20$10$0$0$0$0$0$1000$50$0$1$1$1$1$1$1$0$jpg|gif|png$100$5$0$0$$$$$$$$$$$$$$$$$$$$$$$$$$"
				SetArr=Split(Settings,"$")
				
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.Write "<title>�������</title>"
				.Write "</head>"
				.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>"
		        .Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write " <div class='topdashed sort'>"
				If Flag = "Edit" Then
				 .Write "�޸İ���"
				Else
				 .Write "��Ӱ���"
				End If
	            .Write "</div>"
				.Write "<br>"
				
				.write "<div class=tab-page id=boardpanel>"
				.Write "  <form name=GuestBoardForm method=post action=""?Action=Save"">"
				.Write " <SCRIPT type=text/javascript>"& _
				"   var tabPane1 = new WebFXTabPane( document.getElementById( ""boardpanel"" ), 1 )"& _
				" </SCRIPT>"& _
					 
				" <div class=tab-page id=basic-page>"& _
				"  <H2 class=tab>������Ϣ</H2>"& _
				"	<SCRIPT type=text/javascript>"& _
				"				 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"& _
				"	</SCRIPT>" 
				
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.Write "   <input type=""hidden"" name=""GuestBoardID"" value=""" & GuestBoardID & """>"
				.Write "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>����״̬:</strong></td>"
				.Write "            <td>"
				.write "<input type=""radio"" name=""Locked"" value=""0"" "
				If KS.ChkClng(Locked) = 0 Then .Write (" checked")
				.Write ">"
				.Write "����"
				.Write "  <input type=""radio"" name=""Locked"" value=""1"" "
				If KS.ChkClng(Locked) = 1 Then .Write (" checked")
				.Write ">"
				.Write "����"
				.Write "              </td>"
				.Write "          </tr>"
				
				.Write "     <tr  class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�� �� ��:</strong></td>"
				.Write "             <td>"
				.Write "             <select name='parentid'>"
				.Write "               <option value=0>-��Ϊ������-</option>"
				   Dim RST:Set RST=Conn.Execute("Select ID,BoardName From KS_GuestBoard Where ParentID=0 order by orderid")
				   Do While Not RST.Eof
				     If trim(ParentID)=trim(RST(0)) Then
				     .Write "<option value='" & RST(0) & "' selected>" & RST(1) & "</option>"
					 Else
				     .Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
					 End If
				   RST.MoveNext
				   Loop
				   RST.Close
				   Set RST=Nothing
				.Write "             </select>"           
				.Write "              </td>"
				.Write "          </tr>"
				
				
				.Write "          <tr  class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��������:</strong></td>"
				.Write "             <td>"
				.Write "              <input name=""BoardName"" type=""text"" id=""BoardName"" value=""" & BoardName & """ class=""textbox"" style=""width:60%""> �磬����������������ѯ��</td>"
				 .Write "</tr>"
				 .Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>�������:</strong></td>"
				.Write "  <td>"
				.Write "<textarea name=""Note"" cols='75' rows='6' class=""textbox"" style=""height:150px;width:70%"">" & Note &"</textarea>"
				.Write "            </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��ҳ����:</strong></td>"
				.Write "            <td>"
				.Write "              �б�ҳÿҳ��ʾ<input name=""SetArr(20)"" type=""text""  value=""" & SetArr(20) &""" class=""textbox"" style=""width:50;text-align:center""> ����¼  ����ҳÿҳ��ʾ <input name=""SetArr(21)"" type=""text""  value=""" & SetArr(21) &""" class=""textbox"" style=""width:50;text-align:center""> ���ظ���¼"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>������ʾ���:</strong></td>"
				.Write "            <td><input name=""SetArr(42)"" type=""text""  value=""" & SetArr(42) &""" class=""textbox"" style=""width:50px;text-align:center"">Сʱ�����»ظ���������ʾ<span style='color:red'>New</span>��־,����ʾ������0"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��������:</strong></td>"
				.Write "            <td>"
				.Write "              ���������<input name=""SetArr(27)"" type=""text""  value=""" & SetArr(27) &""" class=""textbox"" style=""width:50;text-align:center""> ���һظ�������<input name=""SetArr(28)"" type=""text""  value=""" & SetArr(28) &""" class=""textbox"" style=""width:50;text-align:center"">¥ʱ�Զ�תΪ����"
				.Write "              </td>"
				.Write "          </tr>"

				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�������:</strong></td>"
				.Write "            <td><input type=""hidden"" name=""omaster"" value=""" & master &""">"
				.Write "              <input name=""Master"" type=""text"" id=""Master"" value=""" & Master &""" class=""textbox"" style=""width:50%""> �����������Ӣ�Ķ��Ÿ���"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�� �� ��:</strong></td>"
				.Write "            <td>"
				.Write "              <input name=""OrderID"" type=""text"" value=""" & OrderID &""" class=""textbox""> ���ԽС������Խǰ��"
				.Write "              </td>"
				.Write "          </tr>"

				.Write "</table>"
				.Write "</div>"
				.Write "<div class=tab-page id=""formset"">"
		        .Write " <H2 class=tab>����&���</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""formset"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�Ƿ������ο�����鿴:</strong></td>"
				.Write "            <td>"
				.write "<input type=""radio"" name=""setarr(0)"" value=""1"" "
				If KS.ChkClng(SetArr(0)) = 1 Then .Write (" checked")
				.Write ">"
				.Write "��"
				.Write "  <input type=""radio"" name=""setarr(0)"" value=""0"" "
				If KS.ChkClng(SetArr(0)) = 0 Then .Write (" checked")
				.Write ">"
				.Write "��"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg' style='color:blue'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��ע���û�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(9)' size=5 value='" & setarr(9) & "'> ���Ӻ�ſ����ڱ����淢������</td>"
				.Write "          </tr>"
				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��������˰���Ļ�Ա��:</strong><br/><font color=blue>�������벻Ҫ��ѡ</font></td>"
				.Write "            <td>"
				.Write KS.GetUserGroup_CheckBox("SetArr(1)",SetArr(1),5)
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�����ڴ˰��淢���Ļ�Ա��:</strong><br/><font color=blue>�������벻Ҫ��ѡ</font></td>"
				.Write "            <td>"
				.Write KS.GetUserGroup_CheckBox("SetArr(2)",SetArr(2),5)
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��֤��Ա:</strong><br/><font color=blue>�������˰���Ļ�Ա,�����������ա�����ֻ����֤��Ա�ſ��Խ��루���أ�</font></td>"
				.Write "            <td><textarea name='setarr(10)' style='width:600px;height:140px'>" & setarr(10) & "</textarea><br/><font color=red>�����֤��Ա������Ӣ�Ķ��Ÿ�������kesion1,kesion2�ȡ�</font>"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>����/�ʽ�����:</strong></td>"
				.Write "            <td>�û����ֱ�����ڵ���<input type='text' style='text-align:center' name='setarr(11)' size=5 value='" & setarr(11) & "'>�����ֲſ��Խ���˰������������<br/>�û��ʽ������ڵ���<input type='text' style='text-align:center' name='setarr(12)' size=5 value='" & setarr(12) & "'>Ԫ�ſ��Խ���˰������������</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>һ��ÿ����Ա��෢����:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(13)' size=5 value='" & setarr(13) & "'>ƪ <span style='color:green'>����������0</span>"
				
				.Write "��������������<input type='text' style='text-align:center' name='setarr(40)' size=5 value='" & setarr(40) & "'>���� <span style='color:green'>����������0</span> �������ʱ��<input type='text' style='text-align:center' name='setarr(41)' size=5 value='" & setarr(41) & "'>�� <span style='color:green'>����������0</span>"
				
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�ظ����ѵ�����:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(14)'"
				If trim(SetArr(14))="0" Then .Write " checked"
				.Write " value='0'>������</label>"
				.Write "            <label><input type='radio' name='setarr(14)'"
				If trim(SetArr(14))="1" Then .Write " checked"
				.Write " value='1'>����</label>"
				
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�༭���ѵ�����:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(29)'"
				If trim(SetArr(29))="0" Then .Write " checked"
				.Write " value='0'>������</label>"
				.Write "            <label><input type='radio' name='setarr(29)'"
				If trim(SetArr(29))="1" Then .Write " checked"
				.Write " value='1'>����</label>"
				
				.Write "              </td>"
				.Write "          </tr>"
				
				
			.Write "    <tr vclass=""tdbg"">"
			.Write "      <td height=""25"" align=""right"" width='125' class=""clefttitle""><strong>�����Ա�ϴ�������</strong></td>"
			 .Write "    <td height=""30""><input onclick=""document.getElementById('fj').style.display='';"" name=""SetArr(36)"" type=""radio"" value=""1"""
			 If SetArr(36)="1" Then .Write " Checked"
			 .Write ">���� <input name=""SetArr(36)"" onclick=""document.getElementById('fj').style.display='none';"" type=""radio"" value=""0"""
			 If SetArr(36)="0" Then .Write " Checked"
			 .Write ">������"
			 If SetArr(36)="1" Then
			  .Write "<div id='fj'>"
			 Else
			  .Write "<div id='fj' style='display:none;'>"
			 End If
			 .Write "<font color=green>�����ϴ����ļ����ͣ�<input name=""SetArr(37)"" type=""text"" value=""" & SetArr(37) &""" size='30'>���������|�߸���<br/>�����ϴ����ļ���С��<input name=""SetArr(38)"" type=""text"" value=""" & SetArr(38) &""" style=""text-align:center"" size='8'>KB<br/>ÿ���ϴ��ļ�������<input name=""SetArr(39)"" type=""text"" value=""" & SetArr(39) &""" style=""text-align:center"" size='8'>��,����������0<br/>"
			  .Write "<strong>����ϴ�����ͼƬ�����Զ�����ˮӡ<input type=""checkbox"" name=""SetArr(43)"" value=""1"""
			 if SetArr(43)="1" then .Write " checked"
			 .Write "/></strong></font><br/>"
			 .Write "<br/><strong>�����ڴ˰汾�ϴ��������û���:</strong>"
			 .Write KS.GetUserGroup_CheckBox("SetArr(17)",SetArr(17),5)
			 .Write "<font color=blue>�������벻Ҫ��ѡ</font></div>"
			 .Write "</td></tr>"
				

				.Write "</table>"
				.Write "</div>"
				
				.Write "<div class=tab-page id=""comments"">"
		        .Write " <H2 class=tab>���ӵ�������</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""comments"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"

				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>ÿҳ��ʾ��������:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(44)' size=5 value='" & setarr(44) & "'>�� <span style='color:green'>�˰汾�����õ������ܣ����0��</span></td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��Ա�����ﵽ:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(45)' size=5 value='" & setarr(45) & "'>�� �ſ��ܶ����ӽ��е��� <span style='color:green'>Ϊ��ֹ�����������������ֻ�дﵽһ�������Ļ�Ա���ܷ������,������������0</span></td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�����������е���:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(46)'"
				If trim(SetArr(46))="0" Then .Write " checked"
				.Write " value='0'>������</label>"
				.Write "            <label><input type='radio' name='setarr(46)'"
				If trim(SetArr(46))="1" Then .Write " checked"
				.Write " value='1'>����</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>����Իظ����е���:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(47)'"
				If trim(SetArr(47))="0" Then .Write " checked"
				.Write " value='0'>������</label>"
				.Write "            <label><input type='radio' name='setarr(47)'"
				If trim(SetArr(47))="1" Then .Write " checked"
				.Write " value='1'>����</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��������Լ�������:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(48)'"
				If trim(SetArr(48))="0" Then .Write " checked"
				.Write " value='0'>������</label>"
				.Write "            <label><input type='radio' name='setarr(48)'"
				If trim(SetArr(48))="1" Then .Write " checked"
				.Write " value='1'>����</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>����Ԥ�ù۵�:</strong></td>"
				.Write "            <td><textarea name=""setarr(49)"" cols=""50"" rows=""3"">" & SetArr(49) & "</textarea>"
				.Write "             <br/><span style='color:green'>��ѡ�����۵�����Ӣ�ġ�,���Ÿ�������""��ͬ,����,����""</span> </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>����������շ�����:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(50)'"
				If trim(SetArr(50))="0" Then .Write " checked"
				.Write " value='0'>������</label>"
				.Write "            <label><input type='radio' name='setarr(50)'"
				If trim(SetArr(50))="1" Then .Write " checked"
				.Write " value='1'>����</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				
                .Write "</table>"
				.Write "</div>"				
				
				.Write "<div class=tab-page id=""scores"">"
		        .Write " <H2 class=tab>��������</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""scores"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"

				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>���ظ������ٴﵽ����:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(15)' size=5 value='" & setarr(15) & "'>������ <span style='color:green'>����û����������������õ���ͻ���ֵ����������,����������0</span></td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�ڴ˰������ظ���������:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(16)' size=5 value='" & setarr(16) & "'>������ <span style='color:green'>24Сʱ���ظ�����ֻ��һ��,����������0</span></td>"
				.Write "          </tr>"
				

				.Write "          <tr class='tdbg'>"
				.Write "            <td colspan='2' height=""25""><strong>������������:</strong></td></tr><tr class='tdbg'><td colspan='2'>"
				%>
				<table width="80%" border="0">
  <tr>
    <td align="center">����</td>
    <td align="center"><strong>��������</strong></td>
    <td align="center"><strong>����ظ�</strong></td>
    <td align="center"><strong>�ö�</strong></td>
    <td align="center"><strong>����</strong></td>
    <td align="center"><strong>��ɾ����</strong></td>
    <td align="center"><strong>��ɾ�ظ�</strong></td>
  </tr>
  <tr>
    <td><strong>����</strong></td>
    <td><input type='text' style='text-align:center' name='setarr(3)' size=5 value='<%=setarr(3)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(4)' size=5 value='<%=setarr(4)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(5)' size=5 value='<%=setarr(5)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(6)' size=5 value='<%=setarr(6)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(7)' size=5 value='<%=setarr(7)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(8)' size=5 value='<%=setarr(8)%>'></td>
  </tr>
  <tr>
    <td><strong>����</strong></td>
    <td><input type='text' style='text-align:center' name='setarr(30)' size=5 value='<%=setarr(30)%>' /></td>
    <td><input type='text' style='text-align:center' name='setarr(31)' size=5 value='<%=setarr(31)%>' /></td>
    <td><input type='text' style='text-align:center' name='setarr(32)' size=5 value='<%=setarr(32)%>'/></td>
    <td><input type='text' style='text-align:center' name='setarr(33)' size=5 value='<%=setarr(33)%>'/></td>
    <td><input type='text' style='text-align:center' name='setarr(34)' size=5 value='<%=setarr(34)%>'/></td>
    <td><input type='text' style='text-align:center' name='setarr(35)' size=5 value='<%=setarr(35)%>'/></td>
  </tr>
</table>

				<%
				.Write "</td></tr>"
                .Write "</table>"
				.Write "</div>"
				
				.Write "<div class=tab-page id=""boardrule"">"
		        .Write " <H2 class=tab>���ð��</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""boardrule"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				 .Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>�� ��:</strong><br/><font color=blue>��������</font></td>"
				.Write "  <td>"
				.Write "<textarea name=""BoardRules"" cols='75' rows='6' class=""textbox"" style=""height:180px;width:70%"">" & BoardRules &"</textarea>"
				%>
				<script src="../editor/ckeditor.js"></script>
				<script type="text/javascript">
                CKEDITOR.replace('BoardRules', {width:"99%",height:"300px",toolbar:"Basic",filebrowserBrowseUrl :"../Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			    </script>				

				<%

				.Write "            </td>"
				.Write "          </tr>"
				.Write "</table>"
				.Write "</div>"
				.Write "<div class=tab-page id=""boardclass"">"
		        .Write " <H2 class=tab>�������</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""boardclass"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>�����������:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(23)'"
				If trim(SetArr(23))="0" Then .Write " checked"
				.Write " value='0'>��</label>"
				.Write "            <label><input type='radio' name='setarr(23)'"
				If trim(SetArr(23))="1" Then .Write " checked"
				.Write " value='1'>��</label>"
				
				.Write " &nbsp;&nbsp;<span style='color:#999999'>�����Ƿ��ڱ��������������๦�ܣ�����Ҫͬʱ�趨��Ӧ�ķ���ѡ��������ñ�����</span><td>"
				.Write " </tr>"
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>�����������:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(24)'"
				If trim(SetArr(24))="0" Then .Write " checked"
				.Write " value='0'>��</label>"
				.Write "            <label><input type='radio' name='setarr(24)'"
				If trim(SetArr(24))="1" Then .Write " checked"
				.Write " value='1'>��</label>"
				
				.Write " &nbsp;&nbsp;<span style='color:#999999'>�Ƿ�ǿ���û�����������ʱ����ѡ�����</span><td>"
				.Write " </tr>"
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>���ǰ׺:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="0" Then .Write " checked"
				.Write " value='0'>����ʾ</label> &nbsp;&nbsp; &nbsp;&nbsp;<span style='color:#999999'>�Ƿ�������ǰ����ʾ���������</span>"
				.Write "            <br/><label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="1" Then .Write " checked"
				.Write " value='1'>ֻ��ʾ����</label>"
				.Write "           <br/> <label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="2" Then .Write " checked"
				.Write " value='2'>ֻ��ʾͼ��</label>"
				
				.Write "<td>"
				.Write " </tr>"
				
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>����������:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(26)'"
				If trim(SetArr(26))="0" Then .Write " checked"
				.Write " value='0'>��</label>"
				.Write "            <label><input type='radio' name='setarr(26)'"
				If trim(SetArr(26))="1" Then .Write " checked"
				.Write " value='1'>��</label>"
				
				.Write " &nbsp;&nbsp;<span style='color:#999999'>�û��Ƿ���԰����������ɸѡ�������</span><td>"
				.Write " </tr>"
				
				.Write "<tr class='tdbg'><td colspan='2'>"
				.Write "<tr class='tdbg'><td colspan='2' class='clefttitle' style='text-align:left;font-weight:bold;height:25px'>�������</td></tr>"
				%>
<script type="text/JavaScript">
	var rowtypedata = [
		[
			[1,'', 'tdbg'],
			[1,'<div style="text-align:center">��</div>', 'tdbg'],
			[1,'<input type="text" size="2" name="categoryorder" value="0" />', 'tdbg'],
			[1,'<input type="text" name="categoryname"  size="30"/>', 'tdbg'],
			[1,'<input type="text" name="categoryicon" size="30"/>', 'tdbg'],
			[1,'', 'tdbg']
		],
	];

var addrowdirect = 0;
function addrow(obj, type) {
	var table = obj.parentNode.parentNode.parentNode.parentNode;
	if(!addrowdirect) {
		var row = table.insertRow(obj.parentNode.parentNode.parentNode.rowIndex);
	} else {
		var row = table.insertRow(obj.parentNode.parentNode.parentNode.rowIndex + 1);
	}
	var typedata = rowtypedata[type];
	for(var i = 0; i <= typedata.length - 1; i++) {
		var cell = row.insertCell(i);
		cell.colSpan = typedata[i][0];
		var tmp = typedata[i][1];
		if(typedata[i][2]) {
			cell.className = typedata[i][2];
		}
		tmp = tmp.replace(/\{(\d+)\}/g, function($1, $2) {return addrow.arguments[parseInt($2) + 1];});
		cell.innerHTML = tmp;
	}
	addrowdirect = 0;
}
</script>

<div id="threadtypes_manage">
<table cellspacing="1" width="80%" cellpadding="1" border="0">
<tr style='font-weight:bold;text-align:center' class="title"><td height='22'>ɾ��</td><td>����</td><td>��ʾ˳��</td><td>��������</td><td>ǰ׺ͼ��</td></tr>
<%
If GuestBoardID<>0 Then
  Dim RS:Set RS=Conn.Execute("Select * From KS_GuestCategory Where BoardID=" & GuestBoardID)
  Do While Not RS.Eof
    Response.Write "<tr><td align=""center""><input type=""hidden"" name=""categoryid"" value=""" &rs("categoryid") & """>"
	Response.Write "<input type=""checkbox"" value=""1"" onclick=""if (this.checked){return(confirm('ȷ��ɾ���÷�����?'))}"" name=""categorydel" & RS("CategoryID") & """>"
	Response.Write "</td><td align=""center""><input type=""checkbox"" value=""1"" name=""categorystatus" & RS("CategoryID") & """ "
	if rs("status")="1" then response.write " checked"
	Response.Write "/>"
	response.write "<td><input type=""text"" size=""2"" name=""categoryorder"" value=""" & rs("orderid") &""" /></td>"
	response.write "<td><input type=""text"" name=""categoryname"" size=""30"" value=""" & rs("categoryname") &""" /></td>"
	response.write "<td><input type=""text"" name=""categoryicon""  size=""30"" value=""" & rs("ico") &""" /></td>"
	response.write "</tr>"
  RS.MoveNext
  Loop
  RS.Close
  Set RS=Nothing
End If
%>


<tr><td colspan="6"><div><img src="images/accept.gif" align="absmiddle"/> <a href="#" onclick="addrow(this, 0)" class="addtr">��ӷ���</a></div></td>
</tr>
</table>
</div>				<%
				.Write "</td></tr>"
				
				
				
                .Write "</table>"
				.Write "</div>"
				
				
								
				.Write "  </form>"
				.Write "</body>"
				.Write "</html>"
				.Write "<script language=""JavaScript"">" & vbCrLf
				.Write "<!--" & vbCrLf
				.Write "function CheckForm()" & vbCrLf
				.Write "{ var form=document.GuestBoardForm;" & vbCrLf
				.Write "  if (form.BoardName.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('�������������!');" & vbCrLf
				.Write "    form.BoardName.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   if (form.Note.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('������������!');" & vbCrLf
				.Write "    form.Note.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "      if (form.OrderID.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('������������!');" & vbCrLf
				.Write "    form.OrderID.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   form.submit();"
				.Write "   return true;"
				.Write "}"
				.Write "//-->"
				.Write "</script>"
			 End With
		  End Sub
		  
		  '����
		  Sub GuestBoardSave()
		    Dim categoryid:categoryid=KS.S("categoryid")&",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			Dim CategoryName:CategoryName=KS.S("CategoryName")
			Dim categoryorder:categoryorder=KS.S("categoryorder")
            Dim categoryicon:categoryicon=KS.S("categoryicon")
			Dim categorystatus:categorystatus=KS.S("categorystatus")
			Dim RS,CategoryNameArr,categoryorderArr,categoryiconArr,categorystatusArr,CategoryIDArr
			
			Dim GuestBoardID, RSObj, SqlStr, BoardName, Note, AddDate, Content, Master,Flag, Page, RSCheck,OrderID,ParentID,BoardRules,Settings,I,Locked
			Set RSObj = Server.CreateObject("Adodb.RecordSet")
			Flag = Request.Form("Flag")
			GuestBoardID = Request("GuestBoardID")
			BoardName = Replace(Replace(Request.Form("BoardName"), """", ""), "'", "")
			Note = Replace(Replace(Request.Form("Note"), """", ""), "'", "")
			Master = Request.Form("Master")
			BoardRules=Request.Form("BoardRules")
			OrderID = KS.ChkClng(KS.G("OrderID"))
			ParentID = KS.Chkclng(Request.Form("ParentID"))
			Locked  = KS.ChkClng(Request.Form("Locked"))
			If BoardName = "" Then Call KS.AlertHistory("�������Ʋ���Ϊ��!", -1)
			If Note = "" Then Call KS.AlertHistory("������ܲ���Ϊ��!", -1)
			
			
			For I=0 To 50
			  If I=0 Then 
			   Settings=Request("setarr(" & i & ")") &"$"
			  Else
			   Settings=Settings  & Request("setarr(" & i & ")")& "$"
			  End If
			Next
			
			Set RSObj = Server.CreateObject("Adodb.Recordset")
			If Flag = "Add" Then
			   RSObj.Open "Select top 1 ID From KS_GuestBoard Where BoardName='" & BoardName & "'", Conn, 1, 1
			   If Not RSObj.EOF Then
				  RSObj.Close
				  Set RSObj = Nothing
				  Response.Write ("<script>alert('�Բ���,�����Ѵ���!');history.back(-1);</script>")
				  Exit Sub
			   Else
				RSObj.Close
				RSObj.Open "SELECT top 1 * FROM KS_GuestBoard Where 1=0", Conn, 1, 3
				RSObj.AddNew
				  RSObj("BoardName") = BoardName
				  RSObj("Note") = Note
				  RSObj("AddDate") = Now
				  RSObj("Master") = Master
				  RSObj("OrderID") =OrderID
				  RSObj("ParentID")=ParentID
				  RSObj("lastpost")="0$" & now & "$��$$$$$"
				  RSObj("TodayNum")=0
				  RSObj("PostNum")=0
				  RSObj("TopicNum")=0
				  RSObj("Locked")=Locked
				  RSObj("BoardRules")=BoardRules
				  RSObj("Settings")=Settings
				RSObj.Update
				GuestBoardID=RSObj("ID")
				 RSObj.Close
			If Not KS.IsNul(CategoryName) Then
			   CategoryNameArr=Split(Replace(CategoryName," ",""),",")
			   categoryorder=split(Replace(categoryorder," ",""),",")
			   categoryiconArr=split(Replace(categoryicon," ",""),",")
			   categorystatusArr=split(Replace(categorystatus," ",""),",")
			   Set RS=Server.CreateObject("ADODB.RECORDSET")
			   For I=0 To Ubound(CategoryNameArr) 
		          RS.Open "Select top 1 * From KS_GuestCategory",conn,1,3
				  RS.AddNew
				    RS("CategoryName")=CategoryNameArr(i)
					RS("OrderID")=KS.ChkClng(categoryorder(i))
					RS("Ico")=trim(categoryiconArr(i))
					RS("Status")=1
					RS("BoardID")=GuestBoardID
				  RS.Update
				  RS.Close
               Next
		   End If
				
				
				
			  End If
			   Set RSObj = Nothing
			   Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
			   Response.Write ("<script> if (confirm('������ӳɹ�!���������?')) {location.href='KS.GuestBoard.asp?Action=Add&parentid=" & ParentID &"';}else{location.href='KS.GuestBoard.asp';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=������� >> <font color=red>���Ա��������</font>';}</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_GuestBoard Where BoardName='" & BoardName & "' And ID<>" & GuestBoardID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 Response.Write ("<script>alert('�Բ���,���������Ѵ���!');history.back(-1);</script>")
				 Exit Sub
			  Else
			   RSObj.Close
			   SqlStr = "SELECT top 1 * FROM KS_GuestBoard Where ID=" & GuestBoardID
			   RSObj.Open SqlStr, Conn, 1, 3
				 RSObj("BoardName") = BoardName
				 RSObj("Note") = Note
				 RSObj("Master") = Master
				 RSObj("OrderID") =OrderID
				 RSObj("Locked")=Locked
				 RSObj("ParentID")=ParentID
				 RSObj("BoardRules")=BoardRules
				 RSObj("Settings")=Settings
			   RSObj.Update
			   RSObj.Close
			   Set RSObj = Nothing
			   
			If Not KS.IsNul(CategoryName) Then
			   CategoryNameArr=Split(CategoryName,",")
			   categoryorder=split(Replace(categoryorder," ","")&",,,,,,,,,,,",",")
			   categoryiconArr=split(Replace(categoryicon," ","")&",,,,,,,,,,,",",")
			   categorystatusArr=split(Replace(categorystatus," ","")&",,,,,,,,,,,",",")
			   categoryIdArr=split(Replace(categoryId," ","")&",,,,,,,,,,,",",")
			   Set RS=Server.CreateObject("ADODB.RECORDSET")
			   For I=0 To Ubound(CategoryNameArr)
			      if KS.ChkClng(categoryIdArr(i))<>0 and KS.ChkClng(KS.S("categorydel"&KS.ChkClng(categoryIdArr(i))))=1 Then
				   Conn.Execute("Delete From KS_GuestCategory Where CategoryID=" & KS.ChkClng(categoryIdArr(i)))
				  Else
					  RS.Open "Select top 1 * From KS_GuestCategory Where CategoryID=" & KS.ChkClng(categoryIdArr(i)),conn,1,3
					  If RS.Eof and RS.Bof Then
					   RS.AddNew
					   RS("Status")=1
					  Else
					   RS("Status")=KS.ChkClng(KS.S("categorystatus" & categoryIdArr(i)))
					  End If
						RS("CategoryName")=trim(CategoryNameArr(i))
						RS("OrderID")=KS.ChkClng(categoryorder(i))
						RS("Ico")=trim(categoryiconArr(i))
						RS("BoardID")=GuestBoardID
					  RS.Update
					  RS.Close
				End If
               Next
		   End If
			   
			  End If
			  Application(KS.SiteSN&"_ClubBoard")=empty
			  Application(KS.SiteSN&"ClubIndex")=empty
			  If trim(lcase(KS.g("omaster")))<>trim(lcase(Master)) Then  UpdateMasterToUser
			  Response.Write ("<script>alert('�����޸ĳɹ�!');location.href='KS.GuestBoard.asp?Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=������� >> <font color=red>���Ա��������</font>';</script>")
			End If
		  End Sub
		  
		   '����KS_User��İ���
		  Sub UpdateMasterToUser()	
			   KS.LoadClubBoard
			   dim node,xml,master,masterarr,i
			   set xml=Application(KS.SiteSN&"_ClubBoard")
			   If IsObject(XML) Then
			     
			    for each node in xml.documentelement.selectnodes("row")
				 if node.selectsinglenode("@master").text<>"" then
					  if master="" then
					   master=node.selectsinglenode("@master").text
					  else
					   master=master& "," & node.selectsinglenode("@master").text
					  end if
				 end if
			    next
			   end if
			   dim rs,newmaster,bzgradeid,admingradeid,superbzgradeid,rsg
			   set rs=server.createobject("adodb.recordset")
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='����'",conn,1,1
				 if not rs.eof then
				  bzgradeid=rs("gradeid")
				 else
				  bzgradeid=0
				 end if
				 rs.close
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='����Ա'",conn,1,1
				 if not rs.eof then
				  admingradeid=rs(0)
				 else
				  admingradeid=0
				 end if
				 rs.close
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='��������'",conn,1,1
				 if not rs.eof then
				  superbzgradeid=rs(0)
				 else
				  superbzgradeid=0
				 end if
				 rs.close
			   if not ks.isnul(master) then
			     masterarr=split(master,",")
				 '�ȸ����û�����̳����ID
				 rs.open "select * from ks_user where ClubSpecialPower=3",conn,1,3
				 do while not rs.eof
				      Set RSG=Conn.Execute("select top 1 GradeID,UserTitle from KS_AskGrade where TypeFlag=1 and Special=0 and ClubPostNum<=" & rs("PostNum") & " And score<=" & rs("Score") & " order by score desc,ClubPostNum Desc")
					  If Not RSG.Eof Then
						   rs("clubgradeid")=rsg(0)
					  else 
					       rsg.close
						   set rsg=conn.execute("select top 1 gradeid from KS_AskGrade where TypeFlag=1 and special=0")
						   if not rsg.eof then
						   rs("clubgradeid")=rsg(0)
						   else
					       rs("clubgradeid")=0
						   end if
					  End If
					  rs.update
					  RSG.Close
				   rs.movenext
				 loop
				 rs.close
				 
				 for i=0 to ubound(masterarr)
				  rs.open "select top 1 * from ks_user where groupid<>1 and username='" & replace(masterarr(i),"'","") & "'",conn,1,3
				  if not rs.eof then
				     if rs("ClubSpecialPower")<>2 then
					   rs("ClubSpecialPower")=3
					   rs("clubgradeid")=bzgradeid
					   rs.update
					 end if
				  end if
				  rs.close
				  if i=0 then 
				   newmaster="'" & masterarr(i) & "'"
				  else
				   newmaster=newmaster & ",'" & masterarr(i) & "'"
				  end if
				 next
				 set rs=nothing
				 conn.execute("update ks_user set ClubSpecialPower=0 where username not in(" & newmaster & ") and ClubSpecialPower<>2 and groupid<>1")
				 
			   end if
				 conn.execute("update ks_user set ClubSpecialPower=1,clubgradeid=" & admingradeid & " where groupid=1")
				 conn.execute("update ks_user set clubgradeid=" & superbzgradeid & " where ClubSpecialPower=2")
				 
          End Sub
		  
		  'ɾ��
		  Sub GuestBoardDel()
		  		 Dim K, GuestBoardID, Page
				 Page = KS.G("Page")
				 GuestBoardID = Trim(KS.G("GuestBoardID"))
				 GuestBoardID = Split(GuestBoardID, ",")
				 For k = LBound(GuestBoardID) To UBound(GuestBoardID)

						Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
					    RS.Open "Select PostTable,id From KS_GuestBook Where BoardID=" & GuestBoardID(k),conn,1,1
						Do While Not RS.Eof
						 Conn.Execute("Delete From " & RS(0) & " Where TopicID=" & RS(1))
						 RS.MoveNext
						Loop
						RS.Close : Set RS=Nothing

					
					Conn.Execute ("Delete From KS_GuestBoard Where ID =" & GuestBoardID(k))
					Conn.Execute ("Delete From KS_GuestBoard Where ParentID =" & GuestBoardID(k))
					Conn.Execute ("Delete From KS_GuestCategory Where BoardID =" & GuestBoardID(k))
					Conn.Execute ("Delete From KS_GuestBook Where BoardID=" & GuestBoardID(k))
				 Next
				 Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
				Response.Write ("<script>location.href='KS.GuestBoard.asp?Page=" & Page & "';</script>")
		  End Sub
		  
		  '��հ�������
		  Sub DelTopic()
		        Dim GuestBoardID:GuestBoardID = KS.ChkClng(KS.G("GuestBoardID"))
		        Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select PostTable,id From KS_GuestBook Where BoardID=" & GuestBoardID,conn,1,1
				Do While Not RS.Eof
					 Conn.Execute("Delete From " & RS(0) & " Where TopicID=" & RS(1))
					 RS.MoveNext
				Loop
				Conn.Execute ("Delete From KS_GuestBook Where BoardID=" & GuestBoardID)
				Conn.Execute("Update KS_GuestBoard Set TodayNum=0,TopicNum=0,PostNum=0,LastPost='0$2010-8-20 15:18:16$��$$$$$' Where ID=" & GuestBoardID)
				RS.Close : Set RS=Nothing
				Response.Write ("<script>alert('��ϲ,�ð��������ѱ����!');location.href='KS.GuestBoard.asp';</script>")
		  End Sub
		  
End Class
%>
 
