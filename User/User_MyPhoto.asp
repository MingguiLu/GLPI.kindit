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
Set KSCls = New Admin_MyPhoto
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MyPhoto
        Private KS,KSUser,ChannelID
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private ComeUrl,SelButton,ReadPoint,ShowStyle,PageNum
		Private F_B_Arr,F_V_Arr,ClassID,Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,PicUrls,Action,I,UserDefineFieldArr,UserDefineFieldValueStr,MapMarker
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
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=2
		If KS.C_S(ChannelID,6)<>2 Then Response.End()
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
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>">我发布的<%=KS.C_S(ChannelID,3)%>(<span class="red"><%=Conn.Execute("Select count(id) from " & KS.C_S(ChannelID,2) &" where Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=1">已审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=0">待审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=2">草 稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="3" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=3">被退稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
         </div>
		<%
		Select Case KS.S("Action")
		  Case "Del"
		   Call KSUser.DelItemInfo(ChannelID,ComeUrl)
		  Case "Add","Edit"
		   Call DoAdd()
		  Case "DoSave"
		   Call DoSave()
		  Case "refresh" Call KSUser.RefreshInfo(KS.C_S(ChannelID,2))
		  Case Else
		   Call PhotoList()
		End Select
	   End Sub
	   
	   Sub PhotoList()
			 CurrentPage = KS.ChkClng(KS.S("page"))
			 If  CurrentPage<=0 Then  CurrentPage=1
                                    
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
									Dim Sql:sql = "select a.*,b.foldername from " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id "& Param &" order by AddDate DESC"

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
 %>
 								  <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="user_myphoto.asp?ChannelID=<%=ChannelID%>&Action=Add"><span style="font-size:14px;color:#ff3300">发布<%=KS.C_S(ChannelID,3)%></span></a></div>
<script src="../ks_inc/jquery.imagePreview.1.0.js"></script>
              <table width="98%" border="0" cellspacing="1" cellpadding="1"  align="center">
                             <%
								Set RS=Server.CreateObject("AdodB.Recordset")
								RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td colspan=4 height=30 align='center' valign=top>没有你要的" & KS.C_S(ChannelID,3) & "!</td></tr>"
								 Else
									totalPut = RS.RecordCount
								    If CurrentPage>1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									End If
									Call showContent
				End If
     %>
                            <table>
							 <tr>
                           <form action="User_MyPhoto.asp?channelid=<%=channelid%>" method="post" name="searchform" id="searchform">
                              <td colspan=4 height="45">
                                    <strong><%=KS.C_S(ChannelID,3)%>搜索：</strong>
                                         <select name="Flag">
                                             <option value="0"><%=F_V_Arr(0)%></option>
                                             <option value="1"><%=F_V_Arr(6)%></option>
                                           </select>
                                           
                                         关键字
                                         <input type="text" name="KeyWord" class="textbox" onclick="if (this.value=='关键字'){this.value=''}" value="关键字" size="20" />
                                         &nbsp;
                                         <input class="button" type="submit" name="submit12" value=" 搜 索 " />
							      </td>
                                    </form>
                                </tr>
                        </table>
		  <%
  End Sub
  
  Sub ShowContent()
     Dim I
    Response.Write "<FORM Action=""User_MyPhoto.asp?ChannelID=" & ChannelID & "&Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
           <tr>
		     <td class="splittd" width="10"><INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID"></td>
             <td class="splittd" width="40" align="center"><div style="cursor:pointer;width:33px;height:33px;border:1px solid #f1f1f1;padding:1px"><a href="<%=RS("PhotoUrl")%>" title="<%=rs("title")%>" target="_blank" class="preview"><img  src="<%=RS("PhotoUrl")%>" width="32" height="32"></a></div>
			 </td>
              <td align="left" class="splittd">
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
              <td align="center" class="splittd">
			     <%If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(3))=1 Then%>
		         <a href="?ChannelID=<%=ChannelID%>&action=refresh&id=<%=rs("id")%>" class="box">刷新</a>
	            <%end if%>
											<%if rs("Verific")<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Action=Edit&id=<%=rs("id")%>&page=<%=CurrentPage%>" class="box">修改</a> <a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除<%=KS.C_S(ChannelID,3)%>吗?'))" class="box">删除</a>
											<%else
												 If KS.C_S(ChannelID,42)=0 Then
												  Response.write "---"
												 Else
												  Response.Write "<a class='box' href='?channelid=" & channelid & "&id=" & rs("id") &"&Action=Edit&&page=" & CurrentPage &"'>修改</a> <a href='#' class='box' disabled>删除</a>"
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
						<table border="0" width="100%" cellpadding="0" cellpadding="0">
								    <tr>
									 <td width="240">
									 <label><INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中所有</label>&nbsp;<button id="btn1" class="pn pnc" onClick="return(confirm('确定删除选中的<%=KS.C_S(ChannelID,3)%>吗?'));" type=submit><strong>删除选定</strong></button> </FORM> 
						              </td>
									  <td align="right">        
								<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
                                      </td>
									 </tr>
						</table>
                  
								<%
  End Sub

 '添加图片
 Sub DoAdd()
 		Call KSUser.InnerLocation("发布" & KS.C_S(ChannelID,3) & "")
		if KS.S("Action")="Edit" Then
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select  top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName &"' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
		     If KS.C_S(ChannelID,42) =0 And RSObj("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
			   RSObj.Close():Set RSObj=Nothing
			   Response.Redirect "../plus/error.asp?action=error&message=" & server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
			 End If
		     ClassID  = RSObj("Tid")
			 Title    = RSObj("Title")
			 KeyWords = RSObj("KeyWords")
			 Author   = RSObj("Author")
			 Origin   = RSObj("Origin")
			 Content  = RSObj("PictureContent")
			 Verific  = RSObj("Verific")
			 ReadPoint= RSObj("ReadPoint")
			 If Verific=3 Then Verific=0
			 PicUrls  = RSObj("PicUrls")
			 PhotoUrl = RSObj("PhotoUrl")
			 ShowStyle= RSObj("ShowStyle")
			 PageNum  = RSObj("PageNum")
			 MapMarker= RSObj("MapMarker")
			 '自定义字段
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
				    UserDefineFieldValueStr=RSObj(UserDefineFieldArr(0,I)) &UnitOption& "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & RSObj(UserDefineFieldArr(0,I)) &UnitOption & "||||"
				  End If
				Next
			  End If
		   End If
		   RSObj.Close:Set RSObj=Nothing
		   Selbutton=KS.C_C(ClassID,1)
		Else
		  Call KSUser.CheckMoney(ChannelID)
		  ClassID=KS.S("ClassID"):Author=KSUser.GetUserInfo("RealName"):PicUrls=""
		  If ClassID="" Then ClassID="0"
		  If ClassID="0" Then
		  SelButton="选择栏目..."
		  Else
		  SelButton=KS.C_C(ClassID,1)
		  End If
		  ReadPoint=0 : ShowStyle=4 : PageNum=10
		End If
		If KS.IsNul(Content) Then Content=" "
			  %>
		<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
                  <form id="myform" action="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform">
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
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
                                <td height="25" align="center"><span><%=F_V_Arr(1)%>：</span></td>
                                <td><% Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) %></td>
                           </tr>
						   
						<%if F_B_Arr(18)="1" Then%> <tr class="tdbg">
								<td height="25" align="center"><span><%=F_V_Arr(18)%>：</span></td>
								<td>经纬度：<input value="<%=MapMarker%>" type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='images/edit_add.gif' align='absmiddle' border='0'>添加电子地图标志</a>
								 <script type="text/javascript">
									  function addMap(){
									  new KesionPopup().PopupCenterIframe('电子地图标注','../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),760,430,'auto');
									  }
									  </script>
								</td>
							  </tr>
							<%end if%>
						   
                           <tr class="tdbg">
                                <td height="25" align="center"><span><%=F_V_Arr(0)%>：</span></td>
                                 <td><input name="Title" class="textbox" type="text" id="Title" value="<%=Title%>" style="width:250px; " maxlength="100" /> <span style="color: #FF0000">*</span></td>
                                </tr>
								<%If F_B_Arr(6)=1 Then%>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(6)%>：</span></td>
                                  <td><input name="KeyWords" class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:220px; " /> <a href="javascript:void(0)" onclick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开</span></td>
                                </tr>
								<%end if%>
								<%If F_B_Arr(7)=1 Then%>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(7)%>：</span></td>
                                        <td height="25"><input class="textbox" name="Author" type="text" id="Author" value="<%=Author%>" style="width:220px; " maxlength="30" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的作者<span></td>
                                </tr>
								<%end if%>
								<%If F_B_Arr(8)=1 Then%>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(8)%>：</span></td>
                                        <td><input class="textbox" name="Origin" type="text" id="Origin" value="<%=Origin%>" style="width:220px; " maxlength="100" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的来源<span></td>
							  </tr>
							  <%End if%>
								<%
							  Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
							  %>
							  <tr class="tdbg">
                                        <td height="35" align="center"><span><%=F_V_Arr(2)%>：</span></td>
                                        <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
									 <tr>
									  <td width="350"><input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:230px;" id='PhotoUrl' maxlength="100" />
                                          
                                          <input class="button" type='button' name='Submit3' value='选择图片...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=4',500,360,window,document.myform.PhotoUrl);" />
								      </td>
									  <td>
									  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=<%=ChannelID%>&Type=Pic' frameborder=0 scrolling=no width='300' height='25'> </iframe>
									  </td>
									 </tr>
									 </table>
										
										
										  <%if KS.S("Action")="Add" Then%>
										  <label><input type='checkbox' name='autothumb' id='autothumb' value='1' checked>使用图集的第一幅图</label><%End If%>
										  </td>
							   </tr>
							   <tr class="tdbg">
							      <td height="35" align="center"><span>显示样式：</span></td>
								  <td><table width='80%'><tr><td>
								  <input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='4'<%If ShowStyle="4" Then response.Write " checked"%>><img src='../images/default/p4.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td>
								  <input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='1'<%If ShowStyle="1" Then response.Write " checked"%>><img src='../images/default/p1.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td>
		   <td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='2'<%If ShowStyle="2" Then Response.Write " checked"%>><img src='../images/default/p2.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='3'<%If ShowStyle="3" Then Response.Write " checked"%>><img src='../images/default/p3.gif'></td></tr></table><div style="margin:5px" id="pagenums"
			<%If ShowStyle="1" or ShowStyle="4" Then Response.Write " style='display:none'"%>
			>每页显示<input type="text" name="pagenum" value="<%=PageNum%>" style="text-align:center;width:30px">张</div>
								  </td>
							   </tr>
							
							  <tr class="tdbg">
                                    <td height="40" align="center" nowrap><span><%=F_V_Arr(4)%>：</span></td>
                                    <td><style type="text/css">
			#thumbnails{background:url(../plus/swfupload/images/albviewbg.gif) no-repeat;_height:expression(document.body.clientHeight > 200? "200px": "auto" );}
			#thumbnails div.thumbshow{text-align:center;margin:2px;padding:2px;width:152px;height:155px;border: dashed 1px #B8B808; background:#FFFFF6;float:left}
			#thumbnails div.thumbshow img{width:130px;height:92px;border:1px solid #CCCC00;padding:1px}

			</style>
			<link href="../plus/swfupload/images/default.css" rel="stylesheet" type="text/css" />
			<script type="text/javascript" src="../ks_inc/kesion.box.js"></script>
			<script type="text/javascript" src="../plus/swfupload/swfupload/swfupload.js"></script>
			<script type="text/javascript" src="../plus/swfupload/js/handlers.js"></script>
<script type="text/javascript">
		var swfu;
		var pid=0;
		function SetAddWater(obj){
		 if (obj.checked){
		 swfu.addPostParam("AddWaterFlag","1");
		 }else{
		 swfu.addPostParam("AddWaterFlag","0");
		 }
        }
		//删除已经上传的图片
		function DelUpFiles(pid)
		{
		 $("#thumbshow"+pid).remove();
		}	
		
		function addImage(bigsrc,smallsrc,text) {
			var newImgDiv = document.createElement("div");
			var delstr = '';
			delstr = '<a href="javascript:DelUpFiles('+pid+')" style="color:#ff6600">[删除]</a>';
			newImgDiv.className = 'thumbshow';
			newImgDiv.id = 'thumbshow'+pid;
			document.getElementById("thumbnails").appendChild(newImgDiv);
			newImgDiv.innerHTML = '<a href="'+bigsrc+'" target="_blank"><span id="show'+pid+'"></span></a>';
			newImgDiv.innerHTML += '<div style="margin-top:10px;text-align:left">'+delstr+' <b>注释：</b><input type="hidden" class="pics" id="pic'+pid+'" value="'+bigsrc+'|'+smallsrc+'"/><input type="text" name="picinfo'+pid+'" value="'+text+'" style="width:148px;" /></div>';
		
			var newImg = document.createElement("img");
			newImg.style.margin = "5px";
		
			document.getElementById("show"+pid).appendChild(newImg);
			if (newImg.filters) {
				try {
					newImg.filters.item("DXImageTransform.Microsoft.Alpha").opacity = 0;
				} catch (e) {
					newImg.style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=' + 0 + ')';
				}
			} else {
				newImg.style.opacity = 0;
			}
		
			newImg.onload = function () {
				fadeIn(newImg, 0);
			};
			newImg.src = smallsrc;
			pid++;
			
		}
	
		window.onload = function () {
			swfu = new SWFUpload({
				// Backend Settings
				upload_url: "swfupload.asp",
				post_params: {"BasicType":<%=KS.C_S(ChannelID,6)%>,"ChannelID":<%=ChannelID%>,"AutoRename":4},

				// File Upload Settings
				file_size_limit : "2 MB",	// 2MB
				file_types : "*.jpg; *.gif; *.png",
				file_types_description : "支持.JPG.gif.png格式的图片,可以多选",
				file_upload_limit : 0,

				// Event Handler Settings - these functions as defined in Handlers.js
				//  The handlers are not part of SWFUpload but are part of my website and control how
				//  my website reacts to the SWFUpload events.
				swfupload_preload_handler : preLoad,
				swfupload_load_failed_handler : loadFailed,
				file_queue_error_handler : fileQueueError,
				file_dialog_complete_handler : fileDialogComplete,
				upload_progress_handler : uploadProgress,
				upload_error_handler : uploadError,
				upload_success_handler : uploadSuccess,
				upload_complete_handler : uploadComplete,

				// Button Settings
				//button_image_url : "../plus/swfupload/images/SmallSpyGlassWithTransperancy_17x18d.png",
				button_placeholder_id : "spanButtonPlaceholder",
				button_width: 195,
				button_height: 22,
				button_text : '<span class="button">本地批量上传(单图限制2 MB)</span>',
				button_text_style : '.button { line-height:22px;font-family: Helvetica, Arial, sans-serif;color:#666666;font-size: 14px; } ',
				button_text_top_padding: 3,
				button_text_left_padding: 0,
				button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
				button_cursor: SWFUpload.CURSOR.HAND,
				
				// Flash Settings
				flash_url : "../plus/swfupload/swfupload/swfupload.swf",
				flash9_url : "../plus/swfuploadswfupload/swfupload_FP9.swf",

				custom_settings : {
					upload_target : "divFileProgressContainer"
				},
				
				// Debug Settings
				debug: false
			});
		};
	</script>
	<script type="text/javascript">
	function OnlineCollect(){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("<div style='text-align:left;padding-left:2px'>网上图片地址</div>","<div style='padding:3px'>带http://开头的远程图片地址,每行一张图片地址:<br/><textarea id='collecthttp' style='width:400px;height:150px'></textarea><br/><input type='button' value='确 定' onclick='ProcessCollect()' class='button'/> <input type='button' value='取 消' class='button' onclick='closeWindow()'/></div>",420);
	}
	function AddTJ(){
	var p=new KesionPopup();
	p.MsgBorder=5;
	p.BgColor='#fff';
	p.ShowBackground=false;
	p.popup("<div style='text-align:left;padding-left:2px'>从上传文件中选择</div>","<div style='padding:3px'><strong>小图地址:</strong><input type='text' name='x1' id='x1'> <input type='button' onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择小图")%>&ChannelID=<%=ChannelID%>',550,290,window,$('#x1')[0]);\" value='选择小图' class='button'/><br/><strong>大图地址:</strong><input type='text' name='x2' id='x2'> <input type='button' onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择小图")%>&ChannelID=<%=ChannelID%>',550,290,window,$('#x2')[0]);\" value='选择大图' class='button'/><br/><strong>简要介绍:</strong><input type='text' name='x3' id='x3'><br/><br/><input type='button' value='加 入' onclick='ProcessAddTj()' class='button'/> <input type='button' value='取 消' class='button' onclick='closeWindow()'/></div>",420);
	}
	function ProcessAddTj(){
	  if ($("#x1").val()==''){
	   alert('请选择一张小图地址!');
	   $("#x1").focus();
	   return false;
	  }
	  if ($("#x2").val()==''){
	   alert('请选择一张大图地址!');
	   $("#x2").focus();
	   return false;
	  }
	  addImage($("#x2").val(),$("#x1").val(),$("#x3").val())
	  $("#x2").val('');
	  $("#x1").val('');
	  $("#x3").val('');
	}
	function ProcessCollect(){
	 var collecthttp=$("#collecthttp").val();
	 if (collecthttp==''){
	   alert('请输入远程图片地址,一行一张地址!');
	   $("#collecthttp").focus();
	   return false;
	 }
	 var carr=collecthttp.split('\n');
	 for(var i=0;i<carr.length;i++){
	   
	   var bigsrc=carr[0];
	   var smallsrc=carr[0];
	   addImage(bigsrc,smallsrc,'')
	 }
	 //$("#collecthttp").empty();
	 closeWindow();
	}
	</script>
	
	<table>
		 <tr>
		  <td>

	    <div class="pn" style="margin: -6px 0px 0 0;">
		 <span id="spanButtonPlaceholder"></span>
		 		
		</div>
		 </td>
		 <td>
		 <button type="button"  class="pn" onClick="OnlineCollect()" style="margin: -6px 0px 0 0;"><strong>网上地址</strong></button>
		 <button type="button"  class="pn" onClick="AddTJ();" style="margin: -6px 0px 0 0;"><strong>图片库...</strong></button>
		 </td>
		 </tr>
		</table>

		<label><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)" checked="checked"/>图片添加水印</label>
		<div id="divFileProgressContainer"></div>
		
	<div id="thumbnails"></div>
			<input type='hidden' name='PicUrls' id='PicUrls' value="<%=PicUrls%>">
									
									
									</td>
                              </tr>
								
							  
								<%If F_B_Arr(9)=1 Then%>
							   <tr class="tdbg">
                                        <td align="center"><%=F_V_Arr(9)%>：<br /></td>
                                        <td align="center">
                                       <textarea style="display:none;" name="Content" id="Content"><%=Server.HTMLEncode(Content)%></textarea>
									   <script type="text/javascript">
										CKEDITOR.replace('Content', {width:"98%",height:"150px",toolbar:"Basic",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
										</script>
									   
									   
									   </td>
                                </tr>
                                <%end if%>
								<%If F_B_Arr(16)=1 Then%>
								<tr class="tdbg">
                                        <td height="25" align="center"><span>阅读<%=KS.Setting(45)%>：</span></td>
                                        <td height="25">
										 <input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6"> <span class="msgtips"><%=KS.Setting(46)%> 如果免费阅读请输入“<font color=red>0</font>”</span></td>
                                </tr>
								<%end if%>
								<tr class="tdbg" <%if KS.S("Action")="Edit" And Verific=1 Then response.write " style='display:none'"%>>
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>状态：</span></td>
                                        <td><input name="Status" type="radio" value="0" <%If Verific=0 Then Response.Write " checked"%> />
投搞
                                          <input name="Status" type="radio" value="2" <%If Verific=2 Then Response.Write " checked"%>/>
草稿</td>
							  </tr>
                               <tr class="tdbg">
                            <td></td>
							<td>
							<button class="pn" id="submit1" type="button" onclick="CheckForm()"><strong>OK, 保 存</strong></button></td>
                              </tr>
</table>
                  </form>




			
			 <script type="text/javascript">
		 	 $(document).ready(function(){
				 IniPicUrl();
			  })
			  
			function IniPicUrl()
			{
			 var PicUrls='<%=replace(PicUrls,vbcrlf,"\t\n")%>';
			  var PicUrlArr=null;
			  if (PicUrls!='')
			   { 
				PicUrlArr=PicUrls.split('|||');
			    for ( var i=1 ;i<PicUrlArr.length+1;i++){ 
			      addImage(PicUrlArr[i-1].split('|')[1],PicUrlArr[i-1].split('|')[2],PicUrlArr[i-1].split('|')[0]);
			    }
			   }
			}
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
				if (document.myform.PhotoUrl.value==''<%if KS.S("Action")="Add" Then response.write " && $('#autothumb').attr('checked')==false"%>)
				{
					alert("请输入<%=KS.C_S(ChannelID,3)%>缩略图！");
					document.myform.PhotoUrl.focus();
					return false;
				}
				<%Call KSUser.ShowUserFieldCheck(ChannelID)%>	
				
				 var picSrcs='';
				  var src='';
				  $("#thumbnails").find(".pics").each(function(){
					 src=$(this).next().val().replace('|||','').replace('|','')+'|'+$(this).val()
					 if(picSrcs==''){
					  picSrcs=src;
					 }else{
					  picSrcs+='|||'+src;
					 }
				  });
				  $('#PicUrls').val(picSrcs);
				if ($('input[name=PicUrls]').val()=='')
				{
				  alert('请输入<%=KS.C_S(ChannelID,3)%>内容!');
				  $('input[name=imgurl1]').focus();
				  return false;
				}
				
                    $('#myform').submit();  
				}
				function CheckClassID()
				{
				 if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
				}
			</script>
			 <%
  End Sub
  
  Sub DoSave()
  				Dim ClassID:ClassID=KS.S("ClassID")
				If KS.ChkClng(KS.C_C(ClassID,20))=0 Then
				 Response.Write "<script>alert('对不起,系统设定不能在此栏目发表,请选择其它栏目!');history.back();</script>":Exit Sub
				 End IF
				Dim Title:Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
				Dim KeyWords:KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				Dim Author:Author=KS.LoseHtml(KS.S("Author"))
				Dim Origin:Origin=KS.LoseHtml(KS.S("Origin"))
				Dim ShowStyle:ShowStyle=KS.ChkClng(KS.S("ShowStyle"))
				Dim PageNum:PageNum=KS.ChkClng(KS.S("PageNum"))
				Dim Content
				Content = KS.FilterIllegalChar(Request.Form("Content"))
				Content=KS.ClearBadChr(content)
				If Content="" Then content=" "
				Dim Verific:Verific=KS.ChkClng(KS.S("Status"))
				Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
				Dim PicUrls:PicUrls=KS.S("PicUrls")
				 If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
				 If KS.ChkClng(KS.S("ID"))<>0 Then
				  If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
				 End If
                 If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '特殊VIP用户无需审核
				 
				Call KSUser.CheckDiyField(ChannelID,UserDefineFieldArr)
				  Dim RSObj
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "名称!');history.back();</script>"
				    Exit Sub
				  End IF
	              If PicUrls="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "!');history.back();</script>"
				    Exit Sub
				  End IF
				 If KS.ChkClng(KS.S("autothumb"))=1 And KS.IsNul(PhotoUrl) Then  PhotoUrl=Split(Split(PicUrls,"|||")(0),"|")(2)
	              If PhotoUrl="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "缩略图!');history.back();</script>"
				    Exit Sub
				  End IF
				If KS.ChkClng(KS.S("ID"))=0 Then
				 Dim Fname,FnameType,TemplateID,WapTemplateID
				 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
				 RSC.Open "select TemplateID,FnameType,FsoType,WapTemplateID from KS_Class Where ID='" & ClassID & "'",conn,1,1
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
				  
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select  top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName & "' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Inputer")=KSUser.UserName
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("AddDate")=Now
				End If
				  RSObj("Title")=Title
				  RSObj("Tid")=ClassID
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("PicUrls")=PicUrls
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Origin")=Origin
				  RSObj("ShowStyle")=ShowStyle
				  RSObj("PageNum")=PageNum
				  RSObj("PictureContent")=Content
				  RSObj("Verific")=Verific
				  RSObj("Comment")=1
				  If F_B_Arr(18)="1" Then	RSObj("MapMarker")=KS.S("MapMark")
				  If F_B_Arr(16)=1 Then
				   RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
				  Call KSUser.AddDiyFieldValue(RSObj,UserDefineFieldArr)
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" And KS.ChkClng(KS.S("ID"))=0 Then
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
				 If KS.ChkClng(KS.S("ID"))=0 Then
				  Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
				  Call KS.FileAssociation(ChannelID,InfoID,PicUrls & PhotoUrl & Content ,0)
				  Call KSUser.AddLog(KSUser.UserName,"在栏目[<a href='" & KS.GetFolderPath(ClassID) & "' target='_blank'>" & KS.C_C(ClassID,1) & "</a>]上传了" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""!",2)
				  KS.Echo "<script>if (confirm('" & KS.C_S(ChannelID,3) & "" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?')){location.href='User_MYPhoto.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';}else{location.href='User_MyPhoto.asp?ChannelID=" & ChannelID &"';}</script>"
				Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Content,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,PicUrls & PhotoUrl & Content ,1)
			     Call KSUser.AddLog(KSUser.UserName,"对" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""做了修改!",2)
				 KS.Echo "<script>alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='User_MyPhoto.asp?ChannelID=" & ChannelID &"';</script>"
				End If
  End Sub
  

End Class
%> 
