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
			{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("window.parent.GuestBoardAdd(0);",'添 加(N)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'全 选(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.GuestBoardControl(1);",'编 辑(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.GuestBoardControl(2);",'删 除(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'刷 新(Z)','disabled');
			}
			function DocDisabledContextMenu()
			{
				DisabledContextMenu('FolderID','GuestBoardID','编 辑(E),删 除(D)','编 辑(E)','','','','')
			}
			function GuestBoardAdd(parentid)
			{
				location.href='KS.GuestBoard.asp?Action=Add&parentid='+parentid;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=版面管理中心 >> <font color=red>添加新版面</font>&ButtonSymbol=GO';
			}
			function EditGuestBoard(id)
			{
				location="KS.GuestBoard.asp?Action=Edit&Page="+Page+"&Flag=Edit&GuestBoardID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=版面管理中心 >> <font color=red>编辑版面</font>&ButtonSymbol=GoSave';
			}
			function DelGuestBoard(id)
			{
			if (confirm('如果有子版面将同时被删除,真的要执行删除操作吗?'))
			 location="KS.GuestBoard.asp?Action=Del&Page="+Page+"&GuestBoardid="+id;
			   SelectedFile='';
			}
			function DelTopic(id){
			if (confirm('执行此操作将清空该版本面的所有主题和回复,此操作不可逆请慎重操作!!!'))
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
					  else alert('一次只能编辑一条版面!')	
					SelectedFile='';
					}
				  else if (op==2)    
					 DelGuestBoard(SelectedFile);
				 }
				else 
				 {
				 if (op==1)
				  alertmsg="编辑";
				 else if(op==2)
				  alertmsg="删除"; 
				 else
				  {
				  WindowReload();
				  alertmsg="操作" 
				  }
				 alert('请选择要'+alertmsg+'的版面');
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
			  .Write "<li class='parent' onclick=""GuestBoardAdd(0);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加版面</span></li>"
			  .Write "<li class='parent' onclick=""GuestBoardControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑版面</span></li>"
			  .Write "<li class='parent' onclick=""GuestBoardControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除版面</span></li>"
			  .Write "</ul>"
			

			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr>"			
			.Write "          <td height=""25"" class=""sort"" align=""center"">版面名称</td>"
			.Write "          <td class=""sort""><div align=""center"">版主</div></td>"
			.Write "          <td align=""center"" class=""sort"">帖子数</td>"
			.Write "          <td width=""50"" class=""sort"" align=""center"">排序</td>"
			.Write "          <td class=""sort"" align=""center"">管理操作</td>"
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
					  .Write "  <td class='splittd' align='center'>主题:<font Color=red>" & RSObj("topicnum") & "</font> 总数:<font Color=red>" & RSObj("postnum") & "</font></td>"
					  .Write "  <td class='splittd' align='center'>" & RSOBJ("OrderID") &"</td>"
					  .Write "  <td class='splittd'> <a href='javascript:GuestBoardAdd(" & rsobj("id") & ")'>添加分版</a> | <a href='javascript:EditGuestBoard(" & rsobj("id") & ")'>修改</a> | <a href='javascript:DelGuestBoard(" & rsobj("id") & ")'>删除</a> </td>"
					  .Write "</tr>"
					  Set RS=Conn.Execute("Select ID,BoardName,master,todaynum,postnum,topicnum,orderid From KS_GuestBoard Where ParentID=" & RSObj("ID") & " Order by orderid")
					  Do While not rs.eof
					  .Write "<tr>"
					  .Write "  <td class='splittd' height='20'> &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;|- <span GuestBoardID='" & RS("ID") & "' ondblclick=""EditGuestBoard(this.GuestBoardID)""><img src='Images/folder/folderopen.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & RS("BoardName") & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' align='center'>&nbsp;" & RS("master") & "&nbsp;</td>"
					  .Write "  <td class='splittd' align='center'>主题:<font Color=red>" & RS("topicnum") & "</font> 总数:<font Color=red>" & RS("postnum") & "</font></td>"
					  .Write "  <td class='splittd' align='center'>" & RS("OrderID") &"</td>"
					  .Write "  <td class='splittd'> <a href='#' disabled>添加分版</a> | <a href='javascript:EditGuestBoard(" & rs("id") & ")'>修改</a> | <a href='javascript:DelGuestBoard(" & rs("id") & ")'>删除</a>  | <a href='javascript:DelTopic(" & rs("id") & ")'>清空</a> </td>"
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
					  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "个", CurrentPage, "Action=" & Action)
				End With
			    Set RS=Nothing
			End Sub
			
			'添加修改版面
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
				   BoardRules="暂无版规" : Locked=0
				End If
				Settings=Settings&"0$0$0$1$1$1$1$1$1$20$$0$0$10$1$0$0$0$1$1$20$10$0$0$0$0$0$1000$50$0$1$1$1$1$1$1$0$jpg|gif|png$100$5$0$0$$$$$$$$$$$$$$$$$$$$$$$$$$"
				SetArr=Split(Settings,"$")
				
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.Write "<title>版面管理</title>"
				.Write "</head>"
				.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>"
		        .Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write " <div class='topdashed sort'>"
				If Flag = "Edit" Then
				 .Write "修改版面"
				Else
				 .Write "添加版面"
				End If
	            .Write "</div>"
				.Write "<br>"
				
				.write "<div class=tab-page id=boardpanel>"
				.Write "  <form name=GuestBoardForm method=post action=""?Action=Save"">"
				.Write " <SCRIPT type=text/javascript>"& _
				"   var tabPane1 = new WebFXTabPane( document.getElementById( ""boardpanel"" ), 1 )"& _
				" </SCRIPT>"& _
					 
				" <div class=tab-page id=basic-page>"& _
				"  <H2 class=tab>基本信息</H2>"& _
				"	<SCRIPT type=text/javascript>"& _
				"				 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"& _
				"	</SCRIPT>" 
				
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.Write "   <input type=""hidden"" name=""GuestBoardID"" value=""" & GuestBoardID & """>"
				.Write "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>版面状态:</strong></td>"
				.Write "            <td>"
				.write "<input type=""radio"" name=""Locked"" value=""0"" "
				If KS.ChkClng(Locked) = 0 Then .Write (" checked")
				.Write ">"
				.Write "开放"
				.Write "  <input type=""radio"" name=""Locked"" value=""1"" "
				If KS.ChkClng(Locked) = 1 Then .Write (" checked")
				.Write ">"
				.Write "锁定"
				.Write "              </td>"
				.Write "          </tr>"
				
				.Write "     <tr  class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>父 版 面:</strong></td>"
				.Write "             <td>"
				.Write "             <select name='parentid'>"
				.Write "               <option value=0>-作为父版面-</option>"
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
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>版面名称:</strong></td>"
				.Write "             <td>"
				.Write "              <input name=""BoardName"" type=""text"" id=""BoardName"" value=""" & BoardName & """ class=""textbox"" style=""width:60%""> 如，技术交流、健康咨询等</td>"
				 .Write "</tr>"
				 .Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>版面介绍:</strong></td>"
				.Write "  <td>"
				.Write "<textarea name=""Note"" cols='75' rows='6' class=""textbox"" style=""height:150px;width:70%"">" & Note &"</textarea>"
				.Write "            </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>分页设置:</strong></td>"
				.Write "            <td>"
				.Write "              列表页每页显示<input name=""SetArr(20)"" type=""text""  value=""" & SetArr(20) &""" class=""textbox"" style=""width:50;text-align:center""> 条记录  帖子页每页显示 <input name=""SetArr(21)"" type=""text""  value=""" & SetArr(21) &""" class=""textbox"" style=""width:50;text-align:center""> 条回复记录"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>新贴显示标记:</strong></td>"
				.Write "            <td><input name=""SetArr(42)"" type=""text""  value=""" & SetArr(42) &""" class=""textbox"" style=""width:50px;text-align:center"">小时内有新回复的帖子显示<span style='color:red'>New</span>标志,不显示请输入0"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>热帖设置:</strong></td>"
				.Write "            <td>"
				.Write "              浏览数大于<input name=""SetArr(27)"" type=""text""  value=""" & SetArr(27) &""" class=""textbox"" style=""width:50;text-align:center""> 次且回复数大于<input name=""SetArr(28)"" type=""text""  value=""" & SetArr(28) &""" class=""textbox"" style=""width:50;text-align:center"">楼时自动转为热帖"
				.Write "              </td>"
				.Write "          </tr>"

				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>版面版主:</strong></td>"
				.Write "            <td><input type=""hidden"" name=""omaster"" value=""" & master &""">"
				.Write "              <input name=""Master"" type=""text"" id=""Master"" value=""" & Master &""" class=""textbox"" style=""width:50%""> 多个版主请用英文逗号隔开"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>排 序 号:</strong></td>"
				.Write "            <td>"
				.Write "              <input name=""OrderID"" type=""text"" value=""" & OrderID &""" class=""textbox""> 序号越小，排在越前面"
				.Write "              </td>"
				.Write "          </tr>"

				.Write "</table>"
				.Write "</div>"
				.Write "<div class=tab-page id=""formset"">"
		        .Write " <H2 class=tab>发帖&浏览</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""formset"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>是否允许游客浏览查看:</strong></td>"
				.Write "            <td>"
				.write "<input type=""radio"" name=""setarr(0)"" value=""1"" "
				If KS.ChkClng(SetArr(0)) = 1 Then .Write (" checked")
				.Write ">"
				.Write "是"
				.Write "  <input type=""radio"" name=""setarr(0)"" value=""0"" "
				If KS.ChkClng(SetArr(0)) = 0 Then .Write (" checked")
				.Write ">"
				.Write "否"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg' style='color:blue'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>新注册用户:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(9)' size=5 value='" & setarr(9) & "'> 分钟后才可以在本版面发布帖子</td>"
				.Write "          </tr>"
				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>允许浏览此版面的会员组:</strong><br/><font color=blue>不限制请不要勾选</font></td>"
				.Write "            <td>"
				.Write KS.GetUserGroup_CheckBox("SetArr(1)",SetArr(1),5)
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>允许在此版面发帖的会员组:</strong><br/><font color=blue>不限制请不要勾选</font></td>"
				.Write "            <td>"
				.Write KS.GetUserGroup_CheckBox("SetArr(2)",SetArr(2),5)
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>认证会员:</strong><br/><font color=blue>允许进入此版面的会员,不限制请留空。否则只有认证会员才可以进入（慎重）</font></td>"
				.Write "            <td><textarea name='setarr(10)' style='width:600px;height:140px'>" & setarr(10) & "</textarea><br/><font color=red>多个认证会员，请用英文逗号隔开，如kesion1,kesion2等。</font>"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>积分/资金限制:</strong></td>"
				.Write "            <td>用户积分必须大于等于<input type='text' style='text-align:center' name='setarr(11)' size=5 value='" & setarr(11) & "'>个积分才可以进入此版面浏览及发帖<br/>用户资金必须大于等于<input type='text' style='text-align:center' name='setarr(12)' size=5 value='" & setarr(12) & "'>元才可以进入此版面浏览及发帖</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>一天每个会员最多发帖数:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(13)' size=5 value='" & setarr(13) & "'>篇 <span style='color:green'>不限制请填0</span>"
				
				.Write "发帖字数不少于<input type='text' style='text-align:center' name='setarr(40)' size=5 value='" & setarr(40) & "'>个字 <span style='color:green'>不限制请填0</span> 发帖间隔时间<input type='text' style='text-align:center' name='setarr(41)' size=5 value='" & setarr(41) & "'>秒 <span style='color:green'>不限制请填0</span>"
				
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>回复自已的帖子:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(14)'"
				If trim(SetArr(14))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(14)'"
				If trim(SetArr(14))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>编辑自已的帖子:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(29)'"
				If trim(SetArr(29))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(29)'"
				If trim(SetArr(29))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				
				.Write "              </td>"
				.Write "          </tr>"
				
				
			.Write "    <tr vclass=""tdbg"">"
			.Write "      <td height=""25"" align=""right"" width='125' class=""clefttitle""><strong>允许会员上传附件：</strong></td>"
			 .Write "    <td height=""30""><input onclick=""document.getElementById('fj').style.display='';"" name=""SetArr(36)"" type=""radio"" value=""1"""
			 If SetArr(36)="1" Then .Write " Checked"
			 .Write ">允许 <input name=""SetArr(36)"" onclick=""document.getElementById('fj').style.display='none';"" type=""radio"" value=""0"""
			 If SetArr(36)="0" Then .Write " Checked"
			 .Write ">不允许"
			 If SetArr(36)="1" Then
			  .Write "<div id='fj'>"
			 Else
			  .Write "<div id='fj' style='display:none;'>"
			 End If
			 .Write "<font color=green>允许上传的文件类型：<input name=""SetArr(37)"" type=""text"" value=""" & SetArr(37) &""" size='30'>多个类型用|线隔开<br/>允许上传的文件大小：<input name=""SetArr(38)"" type=""text"" value=""" & SetArr(38) &""" style=""text-align:center"" size='8'>KB<br/>每天上传文件个数：<input name=""SetArr(39)"" type=""text"" value=""" & SetArr(39) &""" style=""text-align:center"" size='8'>个,不限制请填0<br/>"
			  .Write "<strong>如果上传的是图片，则自动增加水印<input type=""checkbox"" name=""SetArr(43)"" value=""1"""
			 if SetArr(43)="1" then .Write " checked"
			 .Write "/></strong></font><br/>"
			 .Write "<br/><strong>允许在此版本上传附件的用户组:</strong>"
			 .Write KS.GetUserGroup_CheckBox("SetArr(17)",SetArr(17),5)
			 .Write "<font color=blue>不限制请不要勾选</font></div>"
			 .Write "</td></tr>"
				

				.Write "</table>"
				.Write "</div>"
				
				.Write "<div class=tab-page id=""comments"">"
		        .Write " <H2 class=tab>帖子点评设置</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""comments"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"

				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>每页显示点评条数:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(44)' size=5 value='" & setarr(44) & "'>条 <span style='color:green'>此版本不启用点评功能，请填“0”</span></td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>会员威望达到:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(45)' size=5 value='" & setarr(45) & "'>分 才可能对帖子进行点评 <span style='color:green'>为防止恶意点评攻击，建议只有达到一定威望的会员才能发表点评,不限制请输入0</span></td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>允许对主题进行点评:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(46)'"
				If trim(SetArr(46))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(46)'"
				If trim(SetArr(46))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>允许对回复进行点评:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(47)'"
				If trim(SetArr(47))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(47)'"
				If trim(SetArr(47))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>允许点评自己的帖子:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(48)'"
				If trim(SetArr(48))="0" Then .Write " checked"
				.Write " value='0'>不允许</label>"
				.Write "            <label><input type='radio' name='setarr(48)'"
				If trim(SetArr(48))="1" Then .Write " checked"
				.Write " value='1'>允许</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>点评预置观点:</strong></td>"
				.Write "            <td><textarea name=""setarr(49)"" cols=""50"" rows=""3"">" & SetArr(49) & "</textarea>"
				.Write "             <br/><span style='color:green'>可选项，多个观点请用英文“,”号隔开，如""赞同,反对,中立""</span> </td>"
				.Write "          </tr>"				
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>点评算入今日发帖数:</strong></td>"
				.Write "            <td><label><input type='radio' name='setarr(50)'"
				If trim(SetArr(50))="0" Then .Write " checked"
				.Write " value='0'>不计数</label>"
				.Write "            <label><input type='radio' name='setarr(50)'"
				If trim(SetArr(50))="1" Then .Write " checked"
				.Write " value='1'>计数</label>"
				.Write "              </td>"
				.Write "          </tr>"				
				
                .Write "</table>"
				.Write "</div>"				
				
				.Write "<div class=tab-page id=""scores"">"
		        .Write " <H2 class=tab>积分威望</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""scores"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"

				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>下载附件最少达到积分:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(15)' size=5 value='" & setarr(15) & "'>个积分 <span style='color:green'>如果用户积分少于这里设置的最低积分值将不能下载,不限制请填0</span></td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>在此版面下载附件需消耗:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(16)' size=5 value='" & setarr(16) & "'>个积分 <span style='color:green'>24小时内重复下载只扣一次,不限制请填0</span></td>"
				.Write "          </tr>"
				

				.Write "          <tr class='tdbg'>"
				.Write "            <td colspan='2' height=""25""><strong>积分威望设置:</strong></td></tr><tr class='tdbg'><td colspan='2'>"
				%>
				<table width="80%" border="0">
  <tr>
    <td align="center">类型</td>
    <td align="center"><strong>发表主题</strong></td>
    <td align="center"><strong>发表回复</strong></td>
    <td align="center"><strong>置顶</strong></td>
    <td align="center"><strong>精华</strong></td>
    <td align="center"><strong>被删主题</strong></td>
    <td align="center"><strong>被删回复</strong></td>
  </tr>
  <tr>
    <td><strong>积分</strong></td>
    <td><input type='text' style='text-align:center' name='setarr(3)' size=5 value='<%=setarr(3)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(4)' size=5 value='<%=setarr(4)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(5)' size=5 value='<%=setarr(5)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(6)' size=5 value='<%=setarr(6)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(7)' size=5 value='<%=setarr(7)%>'></td>
    <td><input type='text' style='text-align:center' name='setarr(8)' size=5 value='<%=setarr(8)%>'></td>
  </tr>
  <tr>
    <td><strong>威望</strong></td>
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
		        .Write " <H2 class=tab>设置版规</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""boardrule"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				 .Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>版 规:</strong><br/><font color=blue>可以留空</font></td>"
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
		        .Write " <H2 class=tab>主题分类</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""boardclass"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>启用主题分类:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(23)'"
				If trim(SetArr(23))="0" Then .Write " checked"
				.Write " value='0'>否</label>"
				.Write "            <label><input type='radio' name='setarr(23)'"
				If trim(SetArr(23))="1" Then .Write " checked"
				.Write " value='1'>是</label>"
				
				.Write " &nbsp;&nbsp;<span style='color:#999999'>设置是否在本版块启用主题分类功能，您需要同时设定相应的分类选项，才能启用本功能</span><td>"
				.Write " </tr>"
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>发帖必须归类:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(24)'"
				If trim(SetArr(24))="0" Then .Write " checked"
				.Write " value='0'>否</label>"
				.Write "            <label><input type='radio' name='setarr(24)'"
				If trim(SetArr(24))="1" Then .Write " checked"
				.Write " value='1'>是</label>"
				
				.Write " &nbsp;&nbsp;<span style='color:#999999'>是否强制用户发表新主题时必须选择分类</span><td>"
				.Write " </tr>"
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>类别前缀:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="0" Then .Write " checked"
				.Write " value='0'>不显示</label> &nbsp;&nbsp; &nbsp;&nbsp;<span style='color:#999999'>是否在主题前面显示分类的名称</span>"
				.Write "            <br/><label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="1" Then .Write " checked"
				.Write " value='1'>只显示文字</label>"
				.Write "           <br/> <label><input type='radio' name='setarr(25)'"
				If trim(SetArr(25))="2" Then .Write " checked"
				.Write " value='2'>只显示图标</label>"
				
				.Write "<td>"
				.Write " </tr>"
				
				.Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>允许按类别浏览:</strong></td><td>"
				.Write "  <label><input type='radio' name='setarr(26)'"
				If trim(SetArr(26))="0" Then .Write " checked"
				.Write " value='0'>否</label>"
				.Write "            <label><input type='radio' name='setarr(26)'"
				If trim(SetArr(26))="1" Then .Write " checked"
				.Write " value='1'>是</label>"
				
				.Write " &nbsp;&nbsp;<span style='color:#999999'>用户是否可以按照主题分类筛选浏览内容</span><td>"
				.Write " </tr>"
				
				.Write "<tr class='tdbg'><td colspan='2'>"
				.Write "<tr class='tdbg'><td colspan='2' class='clefttitle' style='text-align:left;font-weight:bold;height:25px'>主题分类</td></tr>"
				%>
<script type="text/JavaScript">
	var rowtypedata = [
		[
			[1,'', 'tdbg'],
			[1,'<div style="text-align:center">是</div>', 'tdbg'],
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
<tr style='font-weight:bold;text-align:center' class="title"><td height='22'>删除</td><td>启用</td><td>显示顺序</td><td>分类名称</td><td>前缀图标</td></tr>
<%
If GuestBoardID<>0 Then
  Dim RS:Set RS=Conn.Execute("Select * From KS_GuestCategory Where BoardID=" & GuestBoardID)
  Do While Not RS.Eof
    Response.Write "<tr><td align=""center""><input type=""hidden"" name=""categoryid"" value=""" &rs("categoryid") & """>"
	Response.Write "<input type=""checkbox"" value=""1"" onclick=""if (this.checked){return(confirm('确定删除该分类吗?'))}"" name=""categorydel" & RS("CategoryID") & """>"
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


<tr><td colspan="6"><div><img src="images/accept.gif" align="absmiddle"/> <a href="#" onclick="addrow(this, 0)" class="addtr">添加分类</a></div></td>
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
				.Write "    alert('请输入版面名称!');" & vbCrLf
				.Write "    form.BoardName.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   if (form.Note.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('请输入版面介绍!');" & vbCrLf
				.Write "    form.Note.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "      if (form.OrderID.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('请输入版面序号!');" & vbCrLf
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
		  
		  '保存
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
			If BoardName = "" Then Call KS.AlertHistory("版面名称不能为空!", -1)
			If Note = "" Then Call KS.AlertHistory("版面介绍不能为空!", -1)
			
			
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
				  Response.Write ("<script>alert('对不起,名称已存在!');history.back(-1);</script>")
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
				  RSObj("lastpost")="0$" & now & "$无$$$$$"
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
			   Response.Write ("<script> if (confirm('版面添加成功!继续添加吗?')) {location.href='KS.GuestBoard.asp?Action=Add&parentid=" & ParentID &"';}else{location.href='KS.GuestBoard.asp';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=常规管理 >> <font color=red>留言本版面管理</font>';}</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_GuestBoard Where BoardName='" & BoardName & "' And ID<>" & GuestBoardID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 Response.Write ("<script>alert('对不起,版面名称已存在!');history.back(-1);</script>")
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
			  Response.Write ("<script>alert('版面修改成功!');location.href='KS.GuestBoard.asp?Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=常规管理 >> <font color=red>留言本版面管理</font>';</script>")
			End If
		  End Sub
		  
		   '更新KS_User表的版主
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
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='版主'",conn,1,1
				 if not rs.eof then
				  bzgradeid=rs("gradeid")
				 else
				  bzgradeid=0
				 end if
				 rs.close
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='管理员'",conn,1,1
				 if not rs.eof then
				  admingradeid=rs(0)
				 else
				  admingradeid=0
				 end if
				 rs.close
				 rs.open "select top 1 gradeid from KS_AskGrade where typeflag=1 and UserTitle='超级版主'",conn,1,1
				 if not rs.eof then
				  superbzgradeid=rs(0)
				 else
				  superbzgradeid=0
				 end if
				 rs.close
			   if not ks.isnul(master) then
			     masterarr=split(master,",")
				 '先更新用户在论坛级别ID
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
		  
		  '删除
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
		  
		  '清空版面帖子
		  Sub DelTopic()
		        Dim GuestBoardID:GuestBoardID = KS.ChkClng(KS.G("GuestBoardID"))
		        Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select PostTable,id From KS_GuestBook Where BoardID=" & GuestBoardID,conn,1,1
				Do While Not RS.Eof
					 Conn.Execute("Delete From " & RS(0) & " Where TopicID=" & RS(1))
					 RS.MoveNext
				Loop
				Conn.Execute ("Delete From KS_GuestBook Where BoardID=" & GuestBoardID)
				Conn.Execute("Update KS_GuestBoard Set TodayNum=0,TopicNum=0,PostNum=0,LastPost='0$2010-8-20 15:18:16$无$$$$$' Where ID=" & GuestBoardID)
				RS.Close : Set RS=Nothing
				Response.Write ("<script>alert('恭喜,该版面数据已被清空!');location.href='KS.GuestBoard.asp';</script>")
		  End Sub
		  
End Class
%>
 
