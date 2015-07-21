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
Set KSCls = New User_SoftWare
KSCls.Kesion()
Set KSCls = Nothing

Class User_SoftWare
        Private KS,KSUser,ChannelID,F_B_Arr,F_V_Arr
		Private CurrentPage,totalPut
		Private RS,MaxPerPage,SelButton,ReadPoint
		Private SizeUnit,ClassID,Title,KeyWords,Author,DownLB,DownYY,DownSQ,DownSize,DownPT,YSDZ,ZCDZ,JYMM,Origin,Content,Verific,PhotoUrl,DownUrls,RSObj,ID,AddDate,ComeUrl,CurrentOpStr,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
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
		 IF KS.S("ComeUrl")="" Then
     		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		 Else
     		ComeUrl=KS.S("ComeUrl")
		 End If
			IF Cbool(KSUser.UserLoginChecked)=false Then
			  Response.Write "<script>top.location.href='Login';</script>"
			  Exit Sub
			End If
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=3
		If KS.C_S(ChannelID,6)<>3 Then Response.End()
		if conn.execute("select usertf from ks_channel where channelid=" & channelid)(0)=0 then
		  Response.Write "<script>alert('本频道关闭投稿!');window.close();</script>"
		  Exit Sub
		end if

		'设置缩略图参数
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
		F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
			Call KSUser.Head()
			%>
			<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>">我发布的<%=KS.C_S(ChannelID,3)%>(<span class="red"><%=Conn.Execute("Select count(id) from " & KS.C_S(ChannelID,2) &" where Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&Status=1">已审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&Status=0">待审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&Status=2">草 稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="3" then response.write " class='select'"%>><a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&Status=3">被退稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
		  </div>
			<%
			Action=KS.S("Action")
			Select Case Action
			 Case "Del"  Call KSUser.DelItemInfo(ChannelID,ComeUrl)
			 Case "Add","Edit"  Call DoAdd()
			 Case "AddSave","EditSave"  Call DoSave()
			 Case "refresh" Call KSUser.RefreshInfo(KS.C_S(ChannelID,2))
			 Case Else  Call SoftWareList
			End Select
		End Sub
		Sub SoftWareList()
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where Inputer='"& KSUser.UserName &"'"
                                    Verific=KS.S("Status")
									If Verific="" or not isnumeric(Verific) Then Verific=4
                                    IF Verific<>4 Then 
									   Param= Param & " and Verific=" & Verific
									End If
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
									End if
									If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
									Dim Sql:sql = "select a.*,foldername from " & KS.C_S(ChannelID,2) & " a inner join KS_Class b on a.tid=b.id "& Param &" order by AddDate DESC"

			  					Select Case Verific
								   Case 0 
								    Call KSUser.InnerLocation("待审" & KS.C_S(ChannelID,3) & "列表")
								   Case 1
								    Call KSUser.InnerLocation("已审" & KS.C_S(ChannelID,3) & "列表")
								   Case 2
								   Call KSUser.InnerLocation("草稿" & KS.C_S(ChannelID,3) & "列表")
								   Case 3
								   Call KSUser.InnerLocation("退稿" & KS.C_S(ChannelID,3) & "列表")
                                   Case Else
								    Call KSUser.InnerLocation("所有" & KS.C_S(ChannelID,3) & "列表")
								   End Select
 %> 								  <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="user_mysoftware.asp?ChannelID=<%=ChannelID%>&Action=Add"><span style="font-size:14px;color:#ff3300">发布<%=KS.C_S(ChannelID,3)%></span></a></div>
                <script src="../ks_inc/jquery.imagePreview.1.0.js"></script>
				<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
                          <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  
								  Response.Write "<tr><td class='tdbg' height=30 colspan=6 valign=top>没有你要的" & KS.C_S(ChannelID,3) & "!</td></tr>"
								 
								 Else
									totalPut = RS.RecordCount
								    If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								  		RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									Call showContent
				               End If
     %>                      <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                     <form action="User_MySoftWare.asp?ChannelID=<%=ChannelID%>" method="post" name="searchform">
                                  <td height="45" colspan="6">
										<strong><%=KS.C_S(ChannelID,3)%>搜索：</strong>
										  <select name="Flag">
										   <option value="0">标题</option>
										   <option value="1">关键字</option>
									      </select>
								
										  关键字
										  <input type="text" name="KeyWord" class="textbox" onfocus="if(this.value=='关键字'){this.value=''}" value="关键字" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" 搜 索 ">
							      </td>
								    </form>
                                </tr>
</table>
		  <%
  End Sub
  
  Sub ShowContent()
     Dim I,PhotoUrl
    Response.Write "<FORM Action=""User_MySoftWare.asp?ChannelID=" & ChannelID & "&Action=Del"" name=""myform"" method=""post"">"
     Do While Not RS.Eof
	 If RS("PhotoUrl")<>"" And Not IsNull(RS("PhotoUrl")) Then
	  PhotoUrl=RS("PhotoUrl")
	 Else
	  PhotoUrl="Images/nopic.gif"
	 End If
         %>
           <tr>
						 <td class="splittd" width="10"><INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID"></td>
						 <td class="splittd" width="33"><div style="cursor:hand;text-align:center;width:33px;height:33px;border:1px solid #f1f1f1;padding:1px;"><a href="<%=PhotoUrl%>" target="_blank" title="<%=rs("title")%>" class="preview"><img  src="<%=PhotoUrl%>" width="32" height="32"></a></div>
						 </td>
                         <td  class="splittd" align="left">
						 <div class="ContentTitle"><a href="../item/show.asp?m=<%=ChannelID%>&d=<%=rs("id")%>" target="_blank"><%=trim(RS("title"))%></a></div>
			  
			  <div class="Contenttips">
			            <span>
						 栏目：[<%=RS("FolderName")%>] 发布人：<%=rs("Inputer")%> 发布时间：<%=KS.GetTimeFormat(rs("AddDate"))%>
						 状态：<%Select Case rs("Verific")
											   Case 0
											     Response.Write "<span style=""color:green"">待审</span>"
											   Case 1
											     Response.Write "<span>已审</span>"
                                               Case 2
											     Response.Write "<span style=""color:red"">草稿</span>"
											   Case 3
											     Response.Write "<span style=""color:blue"">退稿</span>"
                                              end select
											  %>
						 </span>
						</div>
					</td>

                     <td class="splittd" align="center">
					   <%If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(3))=1 Then%>
						 <a href="?ChannelID=<%=ChannelID%>&action=refresh&id=<%=rs("id")%>" class="box">刷新</a>
						<%end if%>
											<%if rs("Verific")<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&Action=Edit&id=<%=rs("id")%>&page=<%=CurrentPage%>" class="box">修改</a> <a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除<%=KS.C_S(ChannelID,3)%>吗?'))" class="box">删除</a>
											<%else
											 If KS.C_S(ChannelID,42)=0 Then
											  Response.write "---"
											 Else
											  Response.Write "<a href='?channelid=" & channelid & "&id=" & rs("id") &"&Action=Edit&&page=" & CurrentPage &"' class='box'>修改</a> <a href='#' disabled class='box'>删除</a>"
											 End If
											end if%>
					 </td>
                    </tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
</table>
           <table border="0" width="100%">
        	<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
				 <td><label><INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中所有</label>&nbsp; <button id="btn1" class="pn pnc" onClick="return(confirm('确定删除选中的<%=KS.C_S(ChannelID,3)%>吗?'));" type=submit><strong>删除选定</strong></button> </FORM> 
									 
				 </td>
				 <td align="right">         
	<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
			  </td>
		</tr>
	<%
  End Sub


  '添加
  Sub DoAdd()
          Call KSUser.InnerLocation("发布" & KS.C_S(ChannelID,3) & "")
				ID=KS.ChkClng(KS.S("ID"))
                IF Action="Edit" Then
				   CurrentOpStr=" OK,修改 "
				   Action="EditSave"
				   Dim DownRS:Set DownRS=Server.CreateObject("ADODB.RECORDSET")
				   DownRS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName &"' and ID=" & KS.S("ID"),Conn,1,1
				   IF DownRS.Eof And DownRS.Bof Then
				     call KS.Alert("参数传递出错!",ComeUrl)
					 Exit Sub
				   Else
				 If KS.C_S(ChannelID,42) =0 And DownRS("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
				   DownRS.Close():Set DownRS=Nothing
				   Response.Redirect "../plus/error.asp?action=error&message=" & server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
				 End If
				    Title=DownRS("Title")
				    PhotoUrl=DownRS("PhotoUrl")
				    DownUrls=DownRS("DownUrls")
					DownUrls=split(DownUrls,"|")(2)
				    ClassID=DownRS("TID")
				    KeyWordS=DownRS("KeyWordS")
				    DownLB=DownRS("DownLB")
				    DownYY=DownRS("DownYY")
				    DownSQ=DownRS("DownSQ")
				    DownPT=DownRS("DownPT")
				    YSDZ=DownRS("YSDZ")
				    ZCDZ=DownRS("ZCDZ")
				    JYMM=DownRS("JYMM")
				    Author=DownRS("Author")
				    Origin=DownRS("Origin")
				    Content=DownRS("DownContent")
				    AddDate=DownRS("AddDate")
				    Verific=DownRS("Verific")
					ReadPoint=DownRS("ReadPoint")
                    DownSize=DownRS("DownSize")
					SizeUnit = Right(DownSize, 2)
					DownSize = Replace(DownSize, SizeUnit, "")
					If DownSize = "0" Then DownSize = ""
						UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
						If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
						  Dim UnitOption
						  If UserDefineFieldArr(11,I)="1" Then
						   UnitOption="@" & KS_A_RS_Obj(UserDefineFieldArr(0,I)&"_Unit")
						  Else
						   UnitOption=""
						  End If
						  If UserDefineFieldValueStr="" Then
							UserDefineFieldValueStr=DownRS(UserDefineFieldArr(0,I)) & UnitOption & "||||"
						  Else
							UserDefineFieldValueStr=UserDefineFieldValueStr & DownRS(UserDefineFieldArr(0,I)) & UnitOption & "||||"
						  End If
						Next
					  End If

				   End If
				   SelButton=KS.C_C(ClassID,1)
				   
				   DownRS.Close:Set DownRS=Nothing
				Else
				      Call KSUser.CheckMoney(ChannelID)
					  CurrentOpStr=" OK,添加 ":Action="AddSave":Verific=0:YSDZ="http://":ZCDZ="http://"
					  Author=KSUser.GetUserInfo("RealName")
					  ClassID=KS.S("ClassID")
					  If ClassID="" Then ClassID="0"
					 If ClassID="0" Then
					 SelButton="选择栏目..."
					 Else
					 SelButton=KS.C_C(ClassID,1)
					 End If
					 ReadPoint=0
				End IF

						'取得下载参数
					 Dim I,DownLBList, DownYYList, DownSQList, DownPTList, RSP, DownLBStr, LBArr, YYArr, SQArr, PTArr, DownYYStr, DownSQStr, DownPTStr
					  Set RSP = Server.CreateObject("Adodb.RecordSet")
					  RSP.Open "Select * From KS_DownParam Where ChannelID=" & ChannelID, conn, 1, 1
					  If Not RSP.Eof Then
					   DownLBStr = RSP("DownLB")
					   DownYYStr = RSP("DownYY")
					   DownSQStr = RSP("DownSQ")
					   DownPTStr = RSP("DownPT")
					  End If
					  RSP.Close
					  Set RSP = Nothing
					  '下载类别
					 ' DownLBList="<option value="""" selected> </option>"
					  LBArr = Split(DownLBStr, vbCrLf)
					  For I = 0 To UBound(LBArr)
					   If LBArr(I) = DownLb Then
						DownLBList = DownLBList & "<option value='" & LBArr(I) & "' Selected>" & LBArr(I) & "</option>"
					   Else
						DownLBList = DownLBList & "<option value='" & LBArr(I) & "'>" & LBArr(I) & "</option>"
					   End If
					  Next
					  '下载语言
					   ' DownYYList="<option value="""" selected> </option>"
					  YYArr = Split(DownYYStr, vbCrLf)
					  For I = 0 To UBound(YYArr)
					   If YYArr(I) = DownYY Then
						DownYYList = DownYYList & "<option value='" & YYArr(I) & "' Selected>" & YYArr(I) & "</option>"
					   Else
						DownYYList = DownYYList & "<option value='" & YYArr(I) & "'>" & YYArr(I) & "</option>"
					   End If
					  Next
					'下载授权
					   ' DownSQList="<option value="""" selected> </option>"
					  SQArr = Split(DownSQStr, vbCrLf)
					  For I = 0 To UBound(SQArr)
					   If SQArr(I) = DownSQ Then
						DownSQList = DownSQList & "<option value='" & SQArr(I) & "' Selected>" & SQArr(I) & "</option>"
					   Else
						DownSQList = DownSQList & "<option value='" & SQArr(I) & "'>" & SQArr(I) & "</option>"
					   End If
					  Next
					'下载平台
					  'DownPTList="<option value="""" selected> </option>"
					  PTArr = Split(DownPTStr, vbCrLf)
					  For I = 0 To UBound(PTArr)
						DownPTList = DownPTList & "<a href='javascript:SetDownPT(""" & PTArr(I) & """)'>" & PTArr(I) & "</a>/"
					  Next
					 %>
		<script type="text/javascript" src="../editor/ckeditor.js"></script>
		<script language="javascript">
		        function GetKeyTags(){
				  var text=escape($('#Title').val());
				  if (text!=''){
					  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
					  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
					  function(data){
						$('#KeyWords').val(unescape(data)).attr("disabled",false);
					  });
				  }else{
				   alert('对不起,请先输入标题!');
				  }
				}
				function SetDownPT(addTitle){
					var str=document.myform.DownPT.value;
					if (document.myform.DownPT.value=="") {
						document.myform.DownPT.value=document.myform.DownPT.value+addTitle;
					}else{
						if (str.substr(str.length-1,1)=="/"){
							document.myform.DownPT.value=document.myform.DownPT.value+addTitle;
						}else{
							document.myform.DownPT.value=document.myform.DownPT.value+"/"+addTitle;
						}
					}
					document.myform.DownPT.focus();
				}

				function SetPhotoUrl()
				{
				 if (document.myform.DownUrl.value!='')
				  document.myform.PhotoUrl.value=document.myform.DownUrl.value.split('|')[1];	
				}
				function SetDownUrlByUpLoad(DownUrlStr,FileSize)
				{  $("#DownUrlS").val(DownUrlStr);
				   <%If F_B_Arr(6)=1 Then%>
				    if (FileSize!=0)
					{ 
					  if (FileSize/1024/1024>1)
					  {
					   $("input[name=SizeUnit]")[1].checked=true;
					   document.getElementById('DownSize').value=(FileSize/1024/1024).toFixed(2); 
					  }
					  else{
					  document.getElementById('DownSize').value=(FileSize/1024).toFixed(2);
					  $("input[name=SizeUnit]")[0].checked=true;
					  }
				   }
				  <%end if%>
				var UrlStrArr;
				   UrlStrArr=DownUrlStr.split('|');
				   for (var i=0;i<UrlStrArr.length-1;i++)
				   {
				   var url=UrlStrArr[i]; 
				   if(url!=null&&url!=''){document.myform.DownUrlS.value=url;} 
				  }
				}
				function CheckClassID()
				{
				if (document.myform.ClassID.value=="0" || document.myform.ClassID.value=='') 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
				}
				function CheckForm()
				{   
					if (document.myform.ClassID.value=="0") 
						{
							alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
							//document.myform.ClassID.focus();
							return false;
					 }		
				 if (document.myform.Title.value=="")
					  {
						alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
						document.myform.Title.focus();
						return false;
					  }
					if (document.myform.DownUrlS.value=='')
					{
						alert("请添加<%=KS.C_S(ChannelID,3)%>！");
						document.myform.DownUrlS.focus();
						return false;
					}
					<%Call KSUser.ShowUserFieldCheck(ChannelID)%>
					document.myform.submit();
					return true;
				}
				 
				</script>

				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
                  <form  action="User_mySoftWare.asp?ChannelID=<%=ChannelID%>&Action=<%=Action%>" method="post" name="myform" id="myform">     
				    <input type="hidden" name="ID" value="<%=ID%>">
				    <input type="hidden" name="comeurl" value="<%=ComeUrl%>">
					        <tr class="title">
							 <td colspan=2 align=center>
							 <%IF KS.S("Action")="Edit" Then
							   response.write "修改" & KS.C_S(ChannelID,3)
							   Else
							    response.write "发布" & KS.C_S(ChannelID,3)
							   End iF
							  %>
							 </td>
							</tr>
                             <tr class="tdbg">
                                   <td width="12%" height="25" align="center"><%=F_V_Arr(1)%>：</td>
                                    <td width="88%"><% Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) %>
									</td>
                              </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(0)%>：</td>
                                        <td><input class="textbox" name="Title" type="text" id="Title" value="<%=Title%>" style="width:250px; " maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                                </tr>
						<%if F_B_Arr(10)=1 Then%>	  
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(10)%>：</span></td>
                                  <td><input name="KeyWords" class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:220px; " /> <a href="javascript:void(0)" onclick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开</span></td>

                                </tr>
					   <%end if%>
						<%if F_B_Arr(11)=1 Then%>	  
                                <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(11)%>：</td>
                                        <td height="25"><input class="textbox" name="Author" type="text" id="Author" style="width:250px; " value="<%=Author%>" maxlength="30" /></td>
                                </tr>
					  <%End If%>
						<%if F_B_Arr(12)=1 Then%>	  
                                <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(12)%>：</td>
                                        <td><input class="textbox" name="Origin" value="<%=Origin%>" type="text" id="Origin" style="width:250px; " maxlength="100" /></td>
								</tr>
					  <%end if%>
						<%if F_B_Arr(6)=1 Then%>	  
								 <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(6)%>：</td>
                                        <td>类别:<select name='DownLB'>
		                               <%=DownLBList%>
		                                </select> 语言:<select name='DownYY' size='1'>
		                               <%=DownYYList%>
		                               </select>授权:<select name='DownSQ' size='1'>
		                               <%=DownSQList%></select><%
									 Response.Write "大小:<input type='text' size=4 id='DownSize' name='DownSize' value='" & DownSize & "'>&nbsp;"
									If SizeUnit = "KB" Then
									Response.Write "              <input name=""SizeUnit"" type=""radio"" value=""KB"" checked id=""kb""><label for=""kb"">KB</label> " & vbCrLf
									Response.Write "              <input type=""radio"" name=""SizeUnit"" value=""MB"" id=""mb""><label for=""mb"">MB</label> " & vbCrLf
									Else
									Response.Write "              <input name=""SizeUnit"" type=""radio"" value=""KB""  id=""kb""><label for=""kb"">KB</label> " & vbCrLf
									Response.Write "              <input type=""radio"" name=""SizeUnit"" value=""MB"" checked id=""mb""><label for=""mb"">MB</label> " & vbCrLf
										End If%>                      
		                               </td>
								</tr>
					<%end if%>
						<%if F_B_Arr(7)=1 Then%>	  
								 <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(7)%>：</td>
                                        <td><input class='textbox' type='text' size=70 name='DownPT' value="<%=DownPT%>"><br>
		                               <font color='#808080'>平台选择
		                                <%=DownPTList%></font></td>
				               </tr>
						<%end iF%>
						<%if F_B_Arr(15)=1 Then%>	  
								<tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(15)%>：</td>
                                        <td><input class="textbox" name="YSDZ" type="text" value="<%=YSDZ%>" id="YSDZ" style="width:250px; " maxlength="100" /></td>
                               </tr>
					   <%end if%>
						<%if F_B_Arr(16)=1 Then%>	  
								<tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(16)%>：</td>
                                        <td><input class="textbox" name="ZCDZ" type="text" value="<%=ZCDZ%>" id="ZCDZ" style="width:250px; " maxlength="100" /></td>

								</tr>
					 <%end if%>
						<%if F_B_Arr(17)=1 Then%>	  
								 <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(17)%>：</td>
                                        <td><input class="textbox" name="JYMM" type="text" value="<%=JYMM%>" id="JYMM" style="width:250px; " maxlength="100" /></td>
                              </tr>
						<%end if%>
                             <%
							  Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
							  %> 
						<%if F_B_Arr(8)=1 Then%>	  
						 <tr class="tdbg">
                              <td height="25" align="center"><%=F_V_Arr(8)%>：</td>
                              <td>
									<table width="100%">
									 <tr>
									  <td><input class="textbox"  name="PhotoUrl" value="<%=PhotoUrl%>" type="text" id="PhotoUrl" style="width:230px; " maxlength="100" /></td>
									 <td><input class="button" type='button' name='Submit3' value='选择图片地址...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&amp;pagetitle=<%=Server.URLEncode("选择图片")%>&amp;ChannelID=3',500,360,window,document.myform.PhotoUrl);" /></td>
									<%if F_B_Arr(9)=1 Then%>
									<td>
									 <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=<%=ChannelID%>&Type=Pic' frameborder=0 scrolling=no width='95%' height='25'> </iframe>
									
									</td>
									<%end if%>
								   </tr>
								   </table>
								 
								 
								 </td>
                              </tr>
					   <%end if%>
							  
								
							   <tr class="tdbg">
                                    <td height="25" align="center"><%=KS.C_S(ChannelID,3)%>地址：</td>
                                    <td valign="top">
									
									<table width="100%" border="0">
									 <tr>
									  <td  width="330" nowrap="nowrap"><input type="text" class="textbox" name='DownUrlS' id='DownUrlS' value='<%=DownUrls%>' size="48"> <span style="color: #FF0000">*</span>
                                      </td>
									  <%if F_B_Arr(13)=1 Then%>
									  <td><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=<%=ChannelID%>' frameborder="0" scrolling="no" width='280' height='25'></iframe>
									  </td>
									  <%end if%>
									  </tr>
									 </table>
									 <span class="msgtips">系统提供的上传功能只适合上传比较小的软件（如ASP源代码压缩包）。如果软件比较大（2M以上），请先使用FTP上传，而不要使用系统提供的上传功能，以免上传出错或过度占用服务器的CPU资源。</span>
                                 </td>
								</tr>
							  

                       
						<%if F_B_Arr(14)=1 Then%>	  
								 <tr class="tdbg">
                                        <td align="center"><%=F_V_Arr(14)%>：<br />
                                          </td>
                                        <td align="center"><textarea name="Content" style="display:none"><%=Server.HTMLEncode(Content)%></textarea>
                                       <script type="text/javascript">
										CKEDITOR.replace('Content', {width:"99%",height:"200px",toolbar:"Basic",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
										</script> 
											 
										</td>
                                                  
                                </tr>
						<%end if%>
						<%If F_B_Arr(24)=1 Then%>
						<tr class="tdbg">
                                        <td height="25" align="center"><span>下载<%=KS.Setting(45)%>：</span></td>
                                        <td height="25"><input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6"><%=KS.Setting(46)%> 如果免费下载请输入“<font color=red>0</font>”
										  </td>
                       </tr>
						<%end if%>
						<%if KS.S("Action")="Edit" And Verific=1 Then%>
								<input type="hidden" name="okverific" value="1">
								<input type="hidden" name="verific" value="1">
								<%else%>
						<tr class="tdbg" >
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>状态：</span></td>
                                        <td height="25">
										 <input name="Status" type="radio" value="0" <%If Verific=0 Then Response.Write " checked"%> />
                                          投搞
                                          <input name="Status" type="radio" value="2" <%If Verific=2 Then Response.Write " checked"%>/>
                                          草稿
										  </td>
                                      </tr>
							  <%end if%>	
					<tr class="tdbg">
					        <td></td>
                            <td><button id="btn1" class="pn" type="button" onClick="return CheckForm();"><strong><%=CurrentOpStr%></strong></button></td>

                    </tr>
                  </form>
</table>
		  
		  <%
  End Sub
  Sub DoSave()
                  ID=KS.ChkClng(KS.S("ID"))
				  ClassID=KS.S("ClassID")
				  If KS.ChkClng(KS.C_C(ClassID,20))=0 Then
				 Response.Write "<script>alert('对不起,系统设定不能在此栏目发表,请选择其它栏目!');history.back();</script>":Exit Sub
				 End IF
				  Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
				  KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				  Author=KS.LoseHtml(KS.S("Author"))
				  DownLB=KS.LoseHtml(KS.S("DownLB"))
				  DownYY=KS.LoseHtml(KS.S("DownYY"))
				  DownSQ=KS.LoseHtml(KS.S("DownSQ"))
				  DownSize=KS.S("DownSize")
				  If DownSize = "" Or Not IsNumeric(DownSize) Then DownSize = 0
		          DownSize = DownSize & KS.S("SizeUnit")
				  DownPT=KS.LoseHtml(KS.S("DownPT"))
				  YSDZ=KS.LoseHtml(KS.S("YSDZ"))
				  ZCDZ=KS.LoseHtml(KS.S("ZCDZ"))
				  JYMM=KS.LoseHtml(KS.S("JYMM"))
				  Origin=KS.LoseHtml(KS.S("Origin"))
				  Content = KS.FilterIllegalChar(Request.Form("Content"))
				  If Content="" Then Content=" "
				  Content=KS.ClearBadChr(content)
				  Verific=KS.ChkClng(KS.S("Status"))
				  If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
				 If KS.ChkClng(KS.S("ID"))<>0 and verific=1  Then
					 If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
				 End If
				 if KS.C_S(ChannelID,42)=2 and KS.ChkClng(KS.S("okverific"))=1 Then verific=1
				 If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '特殊VIP用户无需审核
				  PhotoUrl=KS.LoseHtml(KS.S("PhotoUrl"))
				  DownUrls=KS.S("DownUrls")
				  
				Call KSUser.CheckDiyField(ChannelID,UserDefineFieldArr)			  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');history.back();</script>":Exit Sub
				  If Title="" Then  Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "名称!');history.back();</script>":Exit Sub
	              If DownUrls="" Then Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "!');history.back();</script>": Exit Sub
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				    
				 Dim Fname,FnameType,TemplateID,WapTemplateID
				    If ID=0 Then
					 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
					 RSC.Open "select top 1 * from KS_Class Where ID='" & ClassID & "'",conn,1,1
					 if RSC.Eof Then 
					  Response.end
					 Else
					 FnameType=RSC("FnameType")
					 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
					 TemplateID=RSC("TemplateID")
					 WapTemplateID=RSC("WapTemplateID")
					 End If
					 RSC.Close:Set RSC=Nothing
				   End If	 
					RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & ksuser.username & "' and ID=" & ID,Conn,1,3
					If RSObj.Eof Then
						  RSObj.AddNew
						  RSObj("Inputer")=KSUser.UserName
						  RSObj("Hits")=0
						  RSObj("TemplateID")=TemplateID
						  RSObj("WapTemplateID")=WapTemplateID
						  RSObj("Fname")=FName
						  RSObj("AddDate")=Now
						  RSObj("Rank")="★★★"
					End If
					  RSObj("Title")=Title
					  RSObj("TID")=ClassID
					  RSObj("KeyWords")=KeyWords
					  RSObj("Author")=Author
					  RSObj("DownLB")=DownLB
					  RSObj("DownYY")=DownYY
					  RSObj("DownSQ")=DownSQ
					  RSObj("DownSize")=DownSize
					  RSObj("DownPT")=DownPT
					  RSObj("YSDZ")=YSDZ
					  RSObj("ZCDZ")=ZCDZ
					  RSObj("JYMM")=JYMM
					  RSObj("Origin")=Origin
					  RSObj("DownContent")=Content
					  RSObj("PhotoUrl")=PhotoUrl
					  RSObj("DownUrls")="0|下载地址|" & DownUrls
					  RSObj("Verific")=Verific
					  If F_B_Arr(24)=1 Then
					   RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
					  End If
					  Call KSUser.AddDiyFieldValue(RSObj,UserDefineFieldArr)
					  RSObj.Update
					  RSObj.MoveLast
						Dim InfoID:InfoID=RSObj("ID")
						If Left(Ucase(Fname),2)="ID" Then
							RSObj("Fname") = InfoID & FnameType
							RSObj.Update
						End If
						Fname=RSOBj("Fname")
						
						If Verific=1 Then 
							Call KS.SignUserInfoOK(ChannelID,KSUser.UserName,Title,InfoID)
							If KS.C_S(ChannelID,17)=2  and (KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2) Then
							 Dim KSRObj:Set KSRObj=New Refresh
							 Dim DocXML:Set DocXML=KS.RsToXml(RSObj,"row","root")
							 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
							  KSRObj.ModelID=ChannelID
							  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
							  Call KSRObj.RefreshContent()
							  Set KSRobj=Nothing
							End If
						End If
						 RSObj.Close:Set RSObj=Nothing
				 
			 If ID=0 Then
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
				 Call KS.FileAssociation(ChannelID,InfoID,PhotoUrl & Content & DownUrls ,0)
			     Call KSUser.AddLog(KSUser.UserName,"在栏目[<a href='" & KS.GetFolderPath(ClassID) & "' target='_blank'>" & KS.C_C(ClassID,1) & "</a>]发表了" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""!",3)
			     KS.Echo "<script>if (confirm('" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?')){location.href='User_MySoftWare.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';}else{location.href='User_MySoftWare.asp?ChannelID=" & ChannelID & "';}</script>"
			 Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Content,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,PhotoUrl & Content & DownUrls ,1)
			     Call KSUser.AddLog(KSUser.UserName,"对" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""做了修改!",3)
			     KS.Echo "<script>alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='" & ComeUrl & "';</script>"
		    End If	
  End Sub

End Class
%> 
