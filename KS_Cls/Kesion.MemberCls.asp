<!--#include file="Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'-----------------------------------------------------------------------------------------------
'科汛网站管理系统,会员系统函数类
'版本 v7.0
'-----------------------------------------------------------------------------------------------
Class UserCls
			Private KS,I,Node
			'---------定义会员全局变量开始---------------
			Public ID,GroupID,UserName,PassWord,RndPassword,ChargeType
			'---------定义会员全局变量结束---------------
			
			Private Sub Class_Initialize()
			  Set KS=New PublicCls
			End Sub
			Private Sub Class_Terminate()
			 Set KS=Nothing
			End Sub
		   '**************************************************
			'函数名：UserLoginChecked
			'作  用：判断用户是否登录
			'返回值：true或false
			'**************************************************
			Public Function UserLoginChecked()
                'on error resume next
				UserName = KS.R(Trim(KS.C("UserName")))
				PassWord= KS.R(Trim(KS.C("Password")))
				RndPassword=KS.R(Trim(KS.C("RndPassword")))
				IF UserName="" Then
				   UserLoginChecked=false
				   Exit Function
				ElseIf IsObject(Session(KS.SiteSN&"UserInfo")) Then
				   UserLoginChecked=True
				Else
					Dim UserRs
					   Set UserRS=Server.CreateOBject("ADODB.RECORDSET")
					   UserRS.Open "Select top 1 a.*,b.SpaceSize From KS_User a inner join KS_UserGroup b on a.groupid=b.id Where UserName='" & UserName & "' And PassWord='" & PassWord & "'",Conn,1,1
					IF UserRS.Eof And UserRS.Bof Then
					  UserLoginChecked=false
					  Exit Function
					Else
					  If KS.ChkClng(KS.Setting(35))=1 And trim(RndPassword)<>trim(UserRS("RndPassword")) Then
				         UserLoginChecked=false
						 Response.Write ("<script>alert('发现有人正在使用你的账号，你被迫退出！');parent.location.href='" & KS.GetDomain & "User/UserLogout.asp';</script>")
					     Response.end
					  End If
					  
					      '更新活动时间及在线状态
						  If Not KS.IsNul(session("setonlinestatus")) Then
						   Conn.Execute("Update KS_User Set LastLoginTime=" & SQLNowString & " Where UserName='" & UserName & "'")
						  Else
						   Conn.Execute("Update KS_User Set LastLoginTime=" & SQLNowString & ",IsOnline=1 Where UserName='" & UserName & "'")
						  End If
						  
						  '更新其它会员的在线情况
						  If KS.IsNUL(Application("LastUpdateTime")) or (isDate(Application("LastUpdateTime")) and DateDiff("n",Application("LastUpdateTime"),Now)>CLng(KS.Setting(8))) Then
							 Application("LastUpdateTime")=Now
							 Conn.Execute("Update KS_User set isonline=0 WHERE DateDIff(" & DataPart_S &",LastLoginTime," & SQLNowString & ") > "& CLng(KS.Setting(8)) &" * 60")
						  End If
						  
						  Set Session(KS.SiteSN&"UserInfo")=KS.RsToXml(UserRS,"row","")  '写入session
						  
						  UserLoginChecked=true
					End if
					UserRS.Close:Set UserRS=Nothing
			   End IF
			   If IsObject(Session(KS.SiteSN&"UserInfo")) Then
			   Set Node=Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row")
			   GroupID=Node.SelectSingleNode("@groupid").text
			   ChargeType=Node.SelectSingleNode("@chargetype").text
			   End If
			End Function
			
			Function GetUserInfo(ByVal FieldName)
			   If KS.IsNul(UserName) Or KS.IsNul(PassWord) Then Exit Function
			   'If IsObject(Node) Then
			   ' If Not Node Is Nothing Then
			  '   GetUserInfo=Node.SelectSingleNode("@" & lcase(FieldName)).Text
				' Exit Function
			'	End If
			  ' End If
			   If Not IsObject(Session(KS.SiteSN&"UserInfo")) Then UserLoginChecked
			   If IsObject(Session(KS.SiteSN&"UserInfo")) Then
				   Set Node=Session(KS.SiteSN&"UserInfo").DocumentElement.SelectSingleNode("row")
				   If Not Node Is Nothing Then 
					 GetUserInfo=Node.SelectSingleNode("@" & lcase(FieldName)).Text
				   Else
					 GetUserInfo=""
				   End If
			   End If
			End Function


			Public Property Get GetEdays()
					GetEdays = GetUserInfo("Edays")-DateDiff("D",GetUserInfo("BeginDate"),now())
			End Property


			Sub InnerLocation(msg)
				KS.Echo "<script type=""text/javascript"">$('#locationid').html(""" & msg & """);</script>"
			End Sub
		    
			'取得权限
            Function CheckPower(OpType)
			  CheckPower=KS.FoundInArr(KS.U_G(GroupID,"powerlist"),OpType,",")
			End Function
			Sub CheckPowerAndDie(OpType)
			   If UserLoginChecked=false Then
			    KS.Die "<script>top.location.href='Login';</script>"
			   End If
			   If CheckPower(OpType)=false Then
			    KS.ShowError("对不起,你没有此项操作的权限!")
			   End If
			End Sub
			
			'用户上传目录
			Function GetUserFolder(UserName)
			   If KS.HasChinese(UserName) Then
			   Dim Ce:Set Ce=new CtoeCls
			   UserName="[" & Ce.CTOE(KS.R(UserName)) & "]"
			   Set Ce=Nothing
			   End If
			   
			   GetUserFolder=KS.Setting(3)&KS.Setting(91)&"User/" & username & "/"
			End Function
			
           Sub CheckMoney(ChannelID)
		     
			 '检查每个模型每天最多发信息量
			 If KS.ChkCLng(KS.U_S(GroupID,2))<>-1 Then
			   If KS.ChkClng(Conn.Execute("Select Count(*) From " & KS.C_S(ChannelID,2) &" Where inputer='" & username & "' and Year(AddDate)=" & Year(Now) & " and Month(AddDate)=" & Month(Now) & " and Day(AddDate)=" & Day(Now))(0))>=KS.ChkCLng(KS.U_S(GroupID,2)) Then
		         KS.ShowError("对不起,本频道限定每个会员每天只能发布<span style='color=red'>" & KS.ChkCLng(KS.U_S(GroupID,2)) & "</span>条信息!")
			   End If
			 End If
			 
		     If datediff("n",GetUserInfo("RegDate"),now)<KS.ChkClng(KS.C_S(ChannelID,49)) Then
		      KS.ShowError("本频道要求新注册会员" & KS.ChkClng(KS.C_S(ChannelID,49)) & "分钟后才可以发表!")
			 End If
		     If cdbl(KS.C_S(ChannelID,18))<0 And cdbl(GetUserInfo("Money"))<cdbl(abs(KS.C_S(ChannelID,18))) Then
		      KS.ShowError("在本频道发布信息最少需要消费资金" & abs(KS.C_S(ChannelID,18)) & "元,您当前可用资金为" & GetUserInfo("Money") & "元,请充值续费!")
		     End If
		     If cdbl(KS.C_S(ChannelID,19))<0 And cdbl(GetUserInfo("Point"))<cdbl(abs(KS.C_S(ChannelID,19))) Then
		      KS.ShowError("在本频道发布信息最少需要消费" & KS.Setting(45) & abs(KS.C_S(ChannelID,19)) & KS.Setting(46) & ",您当前可用" & KS.Setting(45) & "为" & GetUserInfo("Point") & KS.Setting(46) & ",请充值续费!")
		     End If
		     If cint(KS.C_S(ChannelID,20))<0 And cint(GetUserInfo("Score"))<abs(KS.C_S(ChannelID,20)) Then
		      KS.ShowError("在本频道发布信息最少需要消费积分" & abs(KS.C_S(ChannelID,20)) & "分,您当前可用积分" & GetUserInfo("Score") & "分,请充值续费!")
		     End If
			 
			 '检查未审核信息以判断积分是否够用
			 Dim RS,XML,Node
			 Set RS=Conn.Execute("Select channelid From KS_ItemInfo Where Inputer='"& UserName&"' and verific=0")
			 If Not RS.Eof Then
			    Set XML=KS.RsToXml(rs,"row","")
			 End If
			 RS.Close:Set RS=Nothing
			 If IsObject(XML) Then
			     Dim TotalScore:TotalScore=0
				 Dim TotalPoint:TotalPoint=0
				 Dim TotalMoney:TotalMoney=0
				 Dim Num:Num=0
			    For Each Node In XML.DocumentElement.SelectNodes("row")
				  Dim ModelID:ModelID=Node.SelectSingleNode("@channelid").text
				  Dim Scores:Scores=Cint(KS.C_S(ModelID,20))
				  Dim Points:Points=Cint(KS.C_S(ModelID,19))
				  Dim Moneys:Moneys=Cint(KS.C_S(ModelID,18))
				  If Scores<0 Then
				   TotalScore=TotalScore+Scores
				  End If
				  If Points<0 Then
				   TotalPoint=TotalPoint+Points
				  End If
				  If Moneys<0 Then
				   TotalMoney=TotalMoney+Moneys
				  End If
				  Num=Num+1
				Next
				
				If TotalMoney<0 Then
				  If cint(Abs(TotalMoney)+abs(KS.C_S(ChannelID,18)))>cint(GetUserInfo("Money"))  and KS.C_S(Channelid,18)<0 Then
		           KS.ShowError("在本频道发布信息最少需要消费资金"& abs(KS.C_S(ChannelID,18)) & "元,您的可用资金<font color=#ff6600>" & GetUserInfo("Money") & "</font>元,因在所有投稿栏目中您有<font color=red>" & Num & "</font>篇文档未审核,导致可用资金不足!")
				  End If
				End If
				
				If TotalPoint<0 Then
				  If cint(Abs(TotalPoint)+abs(KS.C_S(ChannelID,19)))>cint(GetUserInfo("Point")) and KS.C_S(Channelid,19)<0 Then
		           KS.ShowError("在本频道发布信息最少需要消费"& KS.Setting(45) & abs(KS.C_S(ChannelID,19)) & KS.Setting(46) & ",您的可用" & KS.Setting(45) & "<font color=#ff6600>" & GetUserInfo("Point") & "</font>" & KS.Setting(46) & ",因在所有投稿栏目中您有<font color=red>" & Num & "</font>篇文档未审核,导致可用" & KS.Setting(45) & "不足!")
				  End If
				End If
				
				If TotalScore<0 Then
				  If cint(Abs(TotalScore)+abs(KS.C_S(Channelid,20)))>cint(GetUserInfo("Score")) and KS.C_S(Channelid,20)<0 Then
		           KS.ShowError("在本频道发布信息最少需要消费积分" & abs(KS.C_S(ChannelID,20)) & "分,您的可用积分<font color=#ff6600>" & GetUserInfo("Score") & "</font>分,因在所有投稿栏目中您有<font color=red>" & Num & "</font>篇文档未审核,导致可用积分不足!")
				  End If
				End If
			 End If
		   End Sub	
		   
		   '用户使用明细
		   Sub UseLogConsum(BasicType,ChannelID,InfoID,Title)
		     Dim Num:Num=KS.ChkClng(KS.U_S(GroupID,11))
		     If Num<>0 Then
				If KS.ChkClng(Conn.Execute("Select Count(1) From KS_LogConsum Where " & InfoID & " not in(select infoid from ks_logconsum Where year(AddDate)=" & year(Now) & " and month(AddDate)=" & month(now) & " and day(AddDate)=" & day(now) &") and year(AddDate)=" & year(Now) & " and month(AddDate)=" & month(now) & " and day(AddDate)=" & day(now) &" and UserName='" &UserName & "' and BasicType=" & BasicType)(0))>=Num Then
				 Select Case BasicType
				   Case 3 KS.Die "<script>alert('系统限制您所在的用户组级别,每人每天只能下载" & Num & "个!');window.close();</script>"
				   Case 4,7 KS.Die "<script>alert('系统限制您所在的用户组级别,每人每天只能观看" & Num & KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3) &"!');window.close();</script>"
				   Case Else
				    KS.Die "<script>alert('系统限制您所在的用户组级别,每人每天只能查看" & Num & KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3)&"!');window.close();</script>"

				 End SELECT
				End If
			 End If
		     dim rs:set rs=server.createobject("adodb.recordset")
			 rs.open "select top 1  * from KS_LogConsum where channelid=" & channelid &" and infoid=" & infoid & " and username='" & username & "'",conn,1,3
			 if rs.eof and rs.bof then
			   rs.addnew
			   rs("basictype")=basictype
			   rs("channelid")=channelid
			   rs("infoid")=infoid
			   rs("title")=title
			   rs("username")=username
			   rs("adddate")=now
			   rs("times")=1
			 else
			   rs("times")=rs("times")+1
			   rs("adddate")=now
			 end if
			  rs.update
			  rs.close:set rs=nothing
		   End Sub
		   
		   '刷新添加时间
		   Sub RefreshInfo(TableName)
		   If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(3))="0" Then
		    KS.AlertHintScript "对不起，本频道没有开通此功能!"
		   End If
		 If KS.ChkClng(KS.U_S(GroupID,12))=0 Then
		   KS.AlertHintScript "对不起，您没有使用此功能的权限，请联系本站管理员!"
		 End If
		   Dim rsf:set rsf=server.createobject("adodb.recordset")
			   rsf.open "select top 1 adddate from [" & TableName & "] where id=" & ks.chkclng(ks.g("id")),conn,1,3
			   if rsf.eof then
			     rsf.close:set rsf=nothing
				   KS.AlertHintScript "参数传递出错！"
			   end if
			   Dim refreshtime:refreshtime=rsf(0)
			   Dim NextTime:NextTime=DateAdd("n",KS.U_S(GroupID,12),refreshtime)
			   if datediff("s",NextTime,now)<1 then
			    rsf.close:set rsf=nothing
                KS.AlertHintScript "对不起，每次刷新间隔" & KS.U_S(GroupID,12) & "分钟，本条信息下次的刷新时间为" & NextTime & "以后!"
			   else
			     rsf(0)=now
				 rsf.update
			   end if
			   rsf.close:set rsf=nothing
			   KS.AlertHintScript "恭喜，刷新成功!"
		End Sub
		
		   
		   '删除模型信息数据
		   Sub DelItemInfo(ChannelID,ComeUrl)
		        Dim ID:ID=KS.S("ID")
				ID=KS.FilterIDs(ID)
				If ID="" Then Call KS.Alert("你没有选中要删除的" & KS.C_S(ChannelID,3) & "!",ComeUrl):Response.End
				Dim RS,DelIDS,DownField
				'判断是不是下载模型
				If KS.C_S(ChannelID,6)=3 Then
				  DownField=",DownUrls"
				End If
				
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				If KS.ChkClng(KS.U_S(GroupID,1))=1 Then
				RS.Open "Select id " & DownField &"  From " & KS.C_S(ChannelID,2) &" Where Inputer='" & UserName & "' And ID In(" & ID & ")",conn,1,3
				Else
				RS.Open "Select id " & DownField &" From " & KS.C_S(ChannelID,2) &" Where Inputer='" & UserName & "' and Verific<>1 And ID In(" & ID & ")",conn,1,3
				End If
				
				Do While Not RS.Eof
				  If DelIds="" Then DelIDs=RS(0)   Else DelIds=DelIds & "," & RS(0)
				  '=======================================删除附件=========================
				  Dim RSD:Set RSD=Server.CreateObject("ADODB.RECORDSET")
				  RSD.Open "Select FileName From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & ID & ")",Conn,1,1
				  Do While Not RSD.Eof
				   if conn.execute("select top 1 filename From KS_UploadFiles Where InfoID not in(" & ID & ") and FileName like '%" & RSD(0) & "%'").eof Then
				   Call KS.DeleteFile(RSD(0))
				   end if
				   RSD.MoveNext
				  Loop
				  RSD.Close
				  conn.Execute ("Delete From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & rs(0) & ")")
				  
				  '下载系统删除下载文件
				  If KS.C_S(ChannelID,6)=3 Then
				    Dim DownUrls:DownUrls=RS(1)
					Dim DownArr,K,DownItemArr,DownUrl
					If Not KS.IsNul(DownUrls) Then
						DownArr=Split(DownUrls,"|||")
						For K=0 To Ubound(DownArr)
						  DownItemArr = Split(DownArr(k),"|")
						  DownUrl = Replace(DownItemArr(2),KS.Setting(2),"")
						  if conn.execute("select top 1 filename From KS_UploadFiles Where InfoID not in(" & ID & ") and FileName like '%" & DownUrl & "%'").eof Then
						  Call KS.DeleteFile(DownUrl)  '删除
						  end if
						Next
					End If
				  End If
				  '============================================================================================================
				  RS.Delete
				  RS.MoveNext
				Loop
				RS.Close:Set RS=Nothing
				If DelIds<>"" Then
				 Call AddLog(UserName,"删除发表的" & KS.C_S(ChannelID,3) & "操作!" & KS.C_S(ChannelID,3) & "ID:" & DelIds,KS.C_S(ChannelID,6))
				End If
				Conn.Execute("Delete From KS_ItemInfo Where Inputer='" & UserName & "' and Verific<>1 and InfoID in(" & ID & ") and channelid=" & ChannelID)
				if ComeUrl="" then
				Response.Redirect("../index.asp")
				else
				Response.Redirect ComeUrl
				end if
		   End Sub
		   		
			'返回专栏选择框
		  Function UserClassOption(TypeID,Sel)
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select ClassID,ClassName From KS_UserClass Where UserName='" & UserName & "' And TypeID="&TypeID,Conn,1,1
			Do While Not RS.Eof
			  If Sel=RS(0) Then
			  UserClassOption=UserClassOption & "<option value=""" & RS(0) & """ selected>" & RS(1) & "</option>"
			  Else
			  UserClassOption=UserClassOption & "<option value=""" & RS(0) & """>" & RS(1) & "</option>"
			  End iF
			  RS.MoveNext
			Loop
			RS.Close:Set RS=Nothing
		  End Function
			
			'返回相应模型的自定义字段名称数组(仅限会员中心调用)
		   Function KS_D_F_Arr(ChannelID)
		      Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
			  KS_RS_Obj.Open "Select FieldName,Title,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,FieldID,EditorType,ShowUnit,UnitOptions,ParentFieldName,MaxLength From KS_Field Where ChannelID=" & ChannelID &" And ShowOnForm=1 And ShowOnUserForm=1 Order By OrderID Asc",Conn,1,1
			 If Not KS_RS_Obj.Eof Then
			  KS_D_F_Arr=KS_RS_Obj.GetRows(-1)
			 Else
			  KS_D_F_Arr=""
			 End If
			 KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		   End Function

		   '取得会员中心信息添加时的自定义字段
		   Function KS_D_F(ChannelID,ByVal UserDefineFieldValueStr)
		      Dim I,K,F_Arr,O_Arr,F_Value,UnitValue,V_Arr
			  Dim O_Text,O_Value,BRStr,O_Len,F_V
			    F_Arr=KS_D_F_Arr(ChannelID)
                If UserDefineFieldValueStr<>"0" And UserDefineFieldValueStr<>""  Then UserDefineFieldValueStr=Split(UserDefineFieldValueStr,"||||")
              If IsArray(F_Arr) Then
				For I=0 To Ubound(F_Arr,2)
				  If F_Arr(13,I)="0" Then
				    KS_D_F=KS_D_F & "<tr  class=""tdbg"" height=""25""><td class=""clefttitle"" align=""center"">" & F_Arr(1,I) & "：</td>"
					KS_D_F=KS_D_F & " <td>"
					If IsArray(UserDefineFieldValueStr) Then
					    F_Value=UserDefineFieldValueStr(I)
					    If F_Arr(11,I)="1" and instr(F_Value,"@")>0 Then
						V_Arr=Split(F_Value,"@")
					    F_Value=V_Arr(0)
					    UnitValue=V_Arr(1)
						End If
					 Else
					   if lcase(F_Arr(4,I))="now" then
					   F_Value=now
					   elseif lcase(F_Arr(4,I))="date" then
					   F_Value=date
					   else
					   F_Value=F_Arr(4,I)
					   end if
					  If Instr(F_Value,"|")<>0 Then
					    F_Value=LFCls.GetSingleFieldValue("select top 1 " & Split(F_Value,"|")(1) & " from " & Split(F_Value,"|")(0) & " where username='" & UserName & "'") 
					   End If
					 End If

				   Select Case F_Arr(3,I)
				     Case 2
				       KS_D_F=KS_D_F & "<textarea style=""width:" & F_Arr(7,i) & "px;height:" & F_Arr(8,i) & "px"" rows=""5"" class=""textbox"" name=""" & F_Arr(0,i) & """ id=""" & F_Arr(0,i) &""">" & F_Value & "</textarea>"
					 Case 3,11
					  If Instr(F_Value,"[#")<>0 then 
					   KS_D_F=KS_D_F & Replace(F_Value,"]","|select]")
					  Else
					   KS_D_F = KS_D_F & GetSelectOption(ChannelID,UserDefineFieldValueStr,F_Arr,F_Arr(3,I),F_Arr(0,i),F_Arr(7,i),F_Arr(5,I),F_Value)
					  End If
					 Case 6
					    If Instr(F_Value,"[#")<>0 then 
					     KS_D_F=KS_D_F & Replace(F_Value,"]","|radio]")
					    Else
					     KS_D_F=KS_D_F & GetRadioOption(F_Arr(0,I),F_Arr(5,I),F_Value)
						End If
					 Case 7
					 If Instr(F_Value,"[#")<>0 then 
					   KS_D_F=KS_D_F & Replace(F_Value,"]","|checkbox]")
					  Else
					   KS_D_F = KS_D_F & GetCheckBoxOption(F_Arr(0,I),F_Arr(5,I),F_Value)
					  End If
					 Case 10
					    If KS.IsNul(F_Value) Then F_Value=" "
					 	KS_D_F=KS_D_F & "<textarea id=""" & F_Arr(0,I) &""" name=""" & F_Arr(0,I) &""">"& Server.HTMLEncode(F_Value) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('" &  F_Arr(0,I) &"', {width:""99%"",height:""" & F_Arr(8,i) & """,toolbar:""" &  F_Arr(10,i) & """,filebrowserBrowseUrl :""../editor/ksplus/SelectUpFiles.asp"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>"

					 Case Else
					   Dim MaxLength:MaxLength=F_Arr(14,i)
					   If Not IsNumerIc(MaxLength)  Or MaxLength="0" Then MaxLength=255
					   KS_D_F=KS_D_F & "<input type=""text"" maxlength=""" & MaxLength &""" class=""textbox"" style=""width:" & F_Arr(7,i) & "px"" name=""" & F_Arr(0,i) & """ value=""" & F_Value & """>"
				   End Select
				   
				   If F_Arr(11,I)="1" Then 
				      If Instr(F_Value,"[#")<>0 then 
					   KS_D_F=KS_D_F & Replace(F_Value,"]","|unit]")
					  Else
					   KS_D_F=KS_D_F & GetUnitOption(F_Arr(0,i),F_Arr(12,i),UnitValue)
					 End If
				   End If
				   
				   If F_Arr(6,I)=1 Then KS_D_F=KS_D_F & "<font color=red> * </font>"
				   KS_D_F=KS_D_F & " <span style=""color:blue;margin-top:5px"">" &  F_Arr(2,I) & "</span>"
				   if F_Arr(3,I)=9 Then KS_D_F=KS_D_F & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & F_Arr(9,I) & "&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				   KS_D_F=KS_D_F & "   </td>"
				   KS_D_F=KS_D_F & "</tr>"
				 End If
			   Next
			End If
		   End Function
		   
		   '单选
		   Function GetRadioOption(FieldName,OptionValue,SelectValue)
		      Dim O_Arr,K,O_Len,F_V,O_Value,O_Text,Str
		      O_Arr=Split(OptionValue,vbcrlf): O_Len=Ubound(O_Arr)
			  For K=0 To O_Len
				 F_V=Split(O_Arr(K),"|")
				 If O_Arr(K)<>"" Then
					If Ubound(F_V)=1 Then
					  O_Value=F_V(0):O_Text=F_V(1)
					Else
					  O_Value=F_V(0):O_Text=F_V(0)
					End If						   
					If trim(SelectValue)=trim(O_Value) Then
						Str=Str & "<label><input type=""radio"" name=""" & FieldName & """ value=""" & O_Value& """ checked>" & O_Text&"</label>"
					Else
						Str=Str & "<label><input type=""radio"" name=""" & FieldName & """ value=""" & O_Value& """>" & O_Text&"</label>"
				    End If
				End If
			 Next
			 GetRadioOption=Str
		   End Function
		   '多选
		   Function GetCheckBoxOption(FieldName,OptionValue,SelectValue)
		    Dim O_Arr,K,O_Len,F_V,O_Value,O_Text,Str
		     O_Arr=Split(OptionValue,vbcrlf): O_Len=Ubound(O_Arr)
			 For K=0 To O_Len
				 F_V=Split(O_Arr(K),"|")
				 If O_Arr(K)<>"" Then
					 If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					 Else
						O_Value=F_V(0):O_Text=F_V(0)
					 End If						   
				     If KS.FoundInArr(trim(SelectValue),trim(O_Value),",")=true Then
						 str=str & "<input type=""checkbox"" name=""" &FieldName& """ value=""" & O_Value& """ checked>" & O_Text
					 Else
						 str=str & "<input type=""checkbox"" name=""" &FieldName& """ value=""" &O_Value& """>" & O_Text
					 End If
				 End If
			Next
			GetCheckBoxOption=str
		   End Function
		   
		   '单位
		   Function GetUnitOption(FieldName,UnitOption,UnitValue)
		      dim str,K
		      str = " <select name=""" & FieldName & "_Unit"" id=""" & FieldName & "_Unit"">"
			  If Not KS.IsNul(UnitOption) Then
				  Dim UnitOptionsArr:UnitOptionsArr=Split(UnitOption,vbcrlf)
				  For K=0 To Ubound(UnitOptionsArr)
					if trim(UnitValue)=trim(UnitOptionsArr(k)) then
					 str=str & "<option value='" & UnitOptionsArr(k) & "' selected>" & UnitOptionsArr(k) & "</option>"
					else
					 str=str & "<option value='" & UnitOptionsArr(k) & "'>" & UnitOptionsArr(k) & "</option>"                 
					end if
				  Next
			 End If
			 str=str & "</select>"
			 GetUnitOption = str
		   End Function
		   '取得下拉及联动选项
		   Function GetSelectOption(ChannelID,UserDefineFieldValueStr,F_Arr,SelectType,FieldName,Width,OptionValue,SelectValue)
		     Dim Str,O_Arr,O_Len,K,F_V,O_Value,O_Text
		       If SelectType=11 Then
					str="<select style=""width:" & Width & """ id=""" & FieldName &""" name=""" &FieldName & """ onchange=""fill" & FieldName &"(this.value)""><option value=''>---请选择---</option>"
	
				Else
				 str= "<select class=""textbox"" style=""width:" & Width & """ id=""" &FieldName &""" name="""& FieldName & """>"
				End If
				O_Arr=Split(OptionValue,vbcrlf): O_Len=Ubound(O_Arr)
				For K=0 To O_Len
				  F_V=Split(O_Arr(K),"|")
				  If O_Arr(K)<>"" Then
					   If Ubound(F_V)=1 Then
				 	    O_Value=F_V(0):O_Text=F_V(1)
					   Else
						O_Value=F_V(0):O_Text=F_V(0)
					   End If						   
					   If trim(SelectValue)=trim(O_Value) Then
						  str=str & "<option value=""" &O_Value& """ selected>" & O_Text & "</option>"
					   Else
						  str=str & "<option value=""" & O_Value& """>" &O_Text & "</option>"
					   End If
				   End If
			  Next
			  str=str & "</select>"
			  '联动菜单
			  If SelectType=11  Then
				Dim JSStr
				str=str &  GetLDMenuStr(ChannelID,F_Arr,UserDefineFieldValueStr,FieldName,JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
			  End If
			  GetSelectOption=str
		   End Function
		  
		   '取得子联动菜单的字段值
		   Function GetFieldValue(F_Arr,UserDefineFieldValueStr,FieldName)
		     Dim I
			 If IsArray(UserDefineFieldValueStr) Then
			      For I=0 To Ubound(F_Arr,2)
				     If Lcase(F_Arr(0,I))=Lcase(FieldName) Then
					   GetFieldValue=UserDefineFieldValueStr(I)
					   Exit Function
					 End If
				  Next
			 End If
		   End Function
		   
		   
		   
		   '取得联动菜单
		   Function GetLDMenuStr(ChannelID,F_Arr,UserDefineFieldValueStr,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str
		     Dim RSL:Set RSL=Conn.Execute("Select Top 1 FieldName,Title,Options,Width From KS_Field Where ChannelID=" & ChannelID & " and ParentFieldName='" & ParentFieldName & "'")
			 If Not RSL.Eof Then
			     Str=Str & " <select name='" & RSL(0) & "' id='" & RSL(0) & "' onchange='fill" & RSL(0) & "(this.value)' style='width:" & RSL(3) & "px'><option value=''>--请选择--</option>"
				 JSStr=JSStr & "var sub" &ParentFieldName & " = new Array();"
				  Options=RSL(2)
				  OArr=Split(Options,Vbcrlf)
				  For I=0 To Ubound(OArr)
				    Varr=Split(OArr(i),"|")
					If Ubound(Varr)=1 Then 
					 V=Varr(0):F=Varr(1)
					Else
					 V=Varr(0)
					 F=Varr(0)
					End If
				    JSStr=JSStr & "sub" & ParentFieldName&"[" & I & "]=new Array('" & V & "','" & F & "')" &vbcrlf
				  Next
				 Str=Str & "</select>"
				 JSStr=JSStr & "function fill"& ParentFieldName&"(v){" &vbcrlf &_
							   "$('#"& RSL(0)&"').empty();" &vbcrlf &_
							   "$('#"& RSL(0)&"').append('<option value="""">--请选择--</option>');" &vbcrlf &_
							   "for (i=0; i<sub" &ParentFieldName&".length; i++){" & vbcrlf &_
							   " if (v==sub" &ParentFieldName&"[i][0]){document.getElementById('" & RSL(0) & "').options[document.getElementById('" & RSL(0) & "').length] = new Option(sub" &ParentFieldName&"[i][1], sub" &ParentFieldName&"[i][1]);}}" & vbcrlf &_
							   "}"
				 Dim DefaultVAL:DefaultVAL=GetFieldValue(F_Arr,UserDefineFieldValueStr,RSL(0))
				 If Not KS.IsNul(DefaultVAL) Then
				   str=str & "<script>$(document).ready(function(){fill"&ParentFieldName&"($('select[name=" &ParentFieldName&"] option:selected').val()); $('#"& RSL(0)&"').val('" & DefaultVAL & "');})</script>" &vbcrlf
				 End If
				 GetLDMenuStr=str & GetLDMenuStr(ChannelID,F_Arr,UserDefineFieldValueStr,RSL(0),JSStr)
			 Else
			     JSStr=JSStr & "function fill" & ParentFieldName &"(v){}"				 
			 End If
			     
		   End Function
		   
		   
		   '根据用户组返回对应模型的可用栏目
			Sub GetClassByGroupID(ByVal ChannelID,ByVal ClassID,Selbutton)
				Dim SQL,K,Node,ClassStr,Pstr,TJ,SpaceStr,Xml
				KS.LoadClassConfig()
				If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
				Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
				If Xml.length=1 Then
				    For Each Node In Xml
If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Then
					  KS.Echo ("<script>alert('对不起,您没有本栏目投稿的权限!');history.back();</script>")  
					Else				   
					  KS.Echo "<font color=red><b>" & Node.SelectSingleNode("@ks1").text & "</b></font>"
				      KS.Echo "<input type='hidden' value='" & Node.SelectSingleNode("@ks0").text & "' name='ClassID' id='ClassID'>"
					End If
				  Next
				  Exit Sub
				End If
				
			    If KS.C_S(ChannelID,41)="3" Then	
				   KS.Echo "<script src=""showclass.asp?channelid=" & ChannelID &"&classid=" & ClassID & """></script>"
				  Exit Sub
				End If

					
				If KS.C_S(ChannelID,41)="0" Then
					KS.Echo "<select onchange=""if ($('#ClassID>option:selected').attr('pubtf')==0){alert('系统设置不能在此栏目下发表!');}"" name='ClassID' id='ClassID' style='width:250px'>"
					For Each Node In Xml
					  If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or (Node.SelectSingleNode("@ks20").text="0" and Node.SelectSingleNode("@ks19").text="0") Then
					  Else
							SpaceStr=""
							TJ=Node.SelectSingleNode("@ks10").text
							If TJ>1 Then
							 For k = 1 To TJ - 1
								SpaceStr = SpaceStr & "──"
							 Next
							End If
							
							If ClassID=Node.SelectSingleNode("@ks0").text Then
								KS.Echo "<option pubtf='" & Node.SelectSingleNode("@ks20").text & "' value='" & Node.SelectSingleNode("@ks0").text & "' selected>" & SpaceStr& Node.SelectSingleNode("@ks1").text & "</option>"
							Else
								KS.Echo "<option pubtf='" & Node.SelectSingleNode("@ks20").text & "' &  value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & "</option>"
							End If
					  End If
					Next
					KS.Echo "</select>"
					Exit Sub
			   Else
				 ClassStr="<input type='button' name='selbutton' id='selbutton' value='" & Selbutton & "' style='height:21px;width:150px;border:0px;background-color: transparent;background-image:url(images/bt.gif);' onClick=""showdiv();"" /><input type='hidden' name='ClassID' id='ClassID' value=" & classid & ">"	
				 %>
				 <script>
				function SelectFolder(Obj){
					var CurrObj;
					if (Obj.ShowFlag=='True')
					{
						ShowOrDisplay(Obj,'none',true);
						Obj.ShowFlag='False';
					}
					else
					{
						ShowOrDisplay(Obj,'',false);
						Obj.ShowFlag='True';
					}
				}
				function ShowOrDisplay(Obj,Flag,Tag)
				{
					for (var i=0;i<document.all.length;i++)
					{
						CurrObj=document.all(i);
						if (CurrObj.ParentID==Obj.TypeID)
						{
							CurrObj.style.display=Flag;
							if (Tag) 
							if (CurrObj.TypeFlag=='Class') ShowOrDisplay(CurrObj.children(0).children(0).children(0).children(0).children(1).children(0),Flag,Tag);
						}
					}
				}
				function showdiv(){
				$("#regtype").toggle();
				$("select").hide();
				}

				function set(element,id,typename){	
				     $("select").show();
					$("#ClassID").val(id);
					$("#selbutton").val(typename);
					$("#regtype").hide();
					for(var i=0 ; i < document.getElementsByName("selclassid").length ; i++ ){
						if(document.getElementsByName("selclassid")[i].checked == true){
							document.getElementsByName("selclassid")[i].checked=false;
							element.checked=true;
						}
					}
				}
				 </script>
				 <%
				 If KS.C_S(ChannelID,41)=1 Then
				  Response.Write "<div class='regtype' id='regtype' style='display:none'>" & GetAllowClass(ChannelID,GroupID)
				 Else
				 response.write "<div class='regtype' id='regtype' style='display:none'><font color=red>提示：灰色的表示不允许发表或您没有权限发表</font>" & ShowClassTree(channelid,groupid)
				 End If	
				 'Response.Write "<iframe src='about:blank' style=""position:absolute; visibility:inherit;top:0px;left:0px;width:310px;height:280px;z-index:-1;filter='progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)';""></iframe></div>"
				 Response.Write "</div>"
			   End If
                Response.Write ClassStr
			End Sub
			
			'显示自定义字段的表单验证
			Public Sub ShowUserFieldCheck(ChannelID)
			    Dim UserDefineFieldArr,I
				UserDefineFieldArr=KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If Cint(UserDefineFieldArr(6,I))=1 Then Response.Write "if ($('input[name=" & UserDefineFieldArr(0,I) & "]').val()==''){alert('" & UserDefineFieldArr(1,I) & "必须填写!');$('input[name=" & UserDefineFieldArr(0,I) & "]').focus();return false;}" & vbcrlf
				 If (Cint(UserDefineFieldArr(3,I))=4 or Cint(UserDefineFieldArr(3,I))=12) Then Response.Write "if ($('input[name=" & UserDefineFieldArr(0,I) &"]').val()!=''&& CheckNumber($('input[name=" & UserDefineFieldArr(0,I) & "]')[0])==false){alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');$('input[name=" & UserDefineFieldArr(0,I) & "]').focus();return false;}"& vbcrlf
				 If Cint(UserDefineFieldArr(3,I))=5 Then Response.Write "if ($('input[name=" & UserDefineFieldArr(0,i) & "]').val()!=''&&is_date($('input[name=" & UserDefineFieldArr(0,i) & "]').val())==false){alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');$('input[name=" & UserDefineFieldArr(0,I) & "]').focus();return false;}" & vbcrlf
				If UserDefineFieldArr(3,I)=8  and UserDefineFieldArr(6,I)=1 Then Response.Write "if (is_email($('input[name=" & UserDefineFieldArr(0,i) & "]').val())==false){alert('" & UserDefineFieldArr(1,I) & "必须填写正确的邮箱!');$('input[name=" & UserDefineFieldArr(0,I) & "]').focus();return false;}" & vbcrlf
				Next
				End If	
		End Sub
		'检查录入
		Sub CheckDiyField(channelid,byref UserDefineFieldArr)
		        Dim I
		        UserDefineFieldArr=KS_D_F_Arr(ChannelID)
			If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If Cint(UserDefineFieldArr(6,I))=1 And KS.IsNul(KS.G(UserDefineFieldArr(0,I))) Then KS.Die "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>"
				 If (Cint(UserDefineFieldArr(3,I))=4 or Cint(UserDefineFieldArr(3,I))=12) And Not KS.IsNul(KS.G(UserDefineFieldArr(0,I))) And Not Isnumeric(KS.G(UserDefineFieldArr(0,I))) Then KS.Die "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>"
				 If Cint(UserDefineFieldArr(3,I))=5 And Not KS.IsNul(KS.G(UserDefineFieldArr(0,I))) And Not IsDate(KS.G(UserDefineFieldArr(0,I))) Then KS.Die "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>"
				 If Cint(UserDefineFieldArr(3,I))=8 And Not KS.IsValidEmail(KS.G(UserDefineFieldArr(0,I))) and Cint(UserDefineFieldArr(6,I))=1 Then KS.Die "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>"
				 
				Next
			End If
		End Sub	
		'更新自定义字段的值
		Sub AddDiyFieldValue(ByRef RS,UserDefineFieldArr)
		      Dim I
		      If IsArray(UserDefineFieldArr) Then
					For I=0 To Ubound(UserDefineFieldArr,2)
						  If (Not KS.IsNul(KS.G(UserDefineFieldArr(0,I))) And (UserDefineFieldArr(3,I)=4 Or UserDefineFieldArr(3,I)=12)) or  (UserDefineFieldArr(3,I)<>4 and UserDefineFieldArr(3,I)<>12) Then
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RS("" & trim(UserDefineFieldArr(0,I)) & "")=Request.Form(trim(UserDefineFieldArr(0,I)))
							else
							 RS("" & trim(UserDefineFieldArr(0,I)) & "")=KS.G(trim(UserDefineFieldArr(0,I)))
							end if
							If KS.ChkClng(UserDefineFieldArr(11,I))=1  Then
							RS("" & trim(UserDefineFieldArr(0,I)) & "_Unit")=KS.G(Trim(UserDefineFieldArr(0,I))&"_Unit")
							End If
						 End If
					Next
			 End If
		 End Sub
			
		 '**************************************************
		'函数名：ShowClassTree
		'作  用：返回允许投稿的目录树。
		'参  数：FolderID ----选择项ID, ChannelID-----返回频道目录树
		'返回值：允许投稿的整棵树
		'**************************************************
		Public Function ShowClassTree(ChannelID,GroupID)
				Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr
				KS.LoadClassConfig()
				If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
				
				TreeStr="<table style=""margin:8px"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
				  SpaceStr=""
				      TreeStr=TreeStr & "<tr ParentID='" & Node.SelectSingleNode("@ks13").text &"'><td>" & vbcrlf
					  TJ=Node.SelectSingleNode("@ks10").text
					  If TJ>1 Then
						 For k = 1 To TJ - 1
							SpaceStr = SpaceStr & "──"
						 Next
						If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or Node.SelectSingleNode("@ks20").text="0"  Then
						 TreeStr=TreeStr& SpaceStr & "<img src='../user/images/doc.gif'><span disabled TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & " <font color=red>[X]</font></a></span>"
						Else
						  TreeStr = TreeStr & SpaceStr & "<img src='../user/images/doc.gif'><span TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & "</a></span><input type='checkbox' id='selclassid' name='selclassid' onclick=""set(this,this.value,'" & Node.SelectSingleNode("@ks1").text & "');"" value='" & Node.SelectSingleNode("@ks0").text & "'>"
						End If
					  Else
					   If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or Node.SelectSingleNode("@ks20").text="0" Then
						 TreeStr=TreeStr & "<img src='../user/images/m_list_22.gif'><span disabled TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & " <font color=red>[X]</font></a></span>"
					   Else
						 TreeStr = TreeStr & "<img src='../user/images/m_list_22.gif'><span TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & "</a></span><input type='checkbox' id='selclassid' name='selclassid' onclick=""set(this,this.value,'" & Node.SelectSingleNode("@ks1").text & "');"" value='" & Node.SelectSingleNode("@ks0").text & "'>"
						End If
					  End If
						TreeStr=TreeStr & vbcrlf & "</td>"&vbcrlf
						TreeStr=TreeStr & "</tr>" & vbcrlf
				Next
		       TreeStr=TreeStr &"</table>"
		       ShowClassTree=TreeStr
		End Function

		
		Function GetAllowClass(ChannelID,GroupID)
				Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr
				KS.LoadClassConfig()
				If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
				
				TreeStr="<table style=""margin:8px"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
				   If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3)) Then
				   Else
					  SpaceStr=""
				      TreeStr=TreeStr & "<tr ParentID='" & Node.SelectSingleNode("@ks13").text &"'><td>" & vbcrlf
					  TJ=Node.SelectSingleNode("@ks10").text
					  If TJ>1 Then
						 For k = 1 To TJ - 1
							SpaceStr = SpaceStr & "──"
						 Next
						  TreeStr = TreeStr & SpaceStr & "<img src='../user/images/doc.gif'><span TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & "</a></span>"
						  If Node.SelectSingleNode("@ks20").text="1" Then
						  	TreeStr = TreeStr &"<input type='checkbox' id='selclassid' name='selclassid' onclick=""set(this,this.value,'" & Node.SelectSingleNode("@ks1").text & "');"" value='" & Node.SelectSingleNode("@ks0").text & "'>"
						  Else
						  TreeStr = TreeStr &"<input type='checkbox' id='selclassid' name='selclassid' disabled>"
						  End If
					  Else
						 TreeStr = TreeStr & "<img src='../user/images/m_list_22.gif'><span TypeID=" & Node.SelectSingleNode("@ks0").text &" ShowFlag='True' onClick='SelectFolder(this);'><a href='#'>" & Node.SelectSingleNode("@ks1").text & "</a></span>"
						 If Node.SelectSingleNode("@ks20").text="1" Then
						 TreeStr =TreeStr & "<input type='checkbox' id='selclassid' name='selclassid' onclick=""set(this,this.value,'" & Node.SelectSingleNode("@ks1").text & "');"" value='" & Node.SelectSingleNode("@ks0").text & "'>"
						 Else
						  TreeStr =TreeStr & "<input type='checkbox' disabled id='selclassid' name='selclassid'>"
						 End If
					  End If
						TreeStr=TreeStr & vbcrlf & "</td>"&vbcrlf
						TreeStr=TreeStr & "</tr>" & vbcrlf
				  End If
				Next
		       TreeStr=TreeStr &"</table>"
		       GetAllowClass=TreeStr
		End Function
		'增加好友动态
		'参数 username 用户 note 备注 ico图标 1评论 2添加文章 0通用
		Sub AddLog(username,note,ico)
		   Dim UserID:UserID=GetUserInfo("userid")
		  If KS.IsNul(UserID) Then UserID=KS.C("UserID")
		  Conn.Execute("Insert Into KS_UserLog([userid],[username],[note],[adddate],[ico]) values(" & KS.ChkClng(GetUserInfo("userid")) & ",'" & UserName & "','" & KS.FilterIllegalChar(replace(note,"'","""")) & "'," & SqlNowString & "," & ico & ")")
		End Sub
			

           '头部
		   Sub Head()
		   %>
			<div  class="title" style="height:30px;line-height:30px;padding-left:6px"><a href="<%=KS.GetDomain%>" target="_parent">网站首页</a> >> <a href="<%=KS.GetDomain%>user/index.asp">会员中心</a> >> <span class="shadow" id="locationid"></span>  </div>
		   <%
		   End Sub
End Class
%> 
