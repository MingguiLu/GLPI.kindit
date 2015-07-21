<% Option Explicit %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/UploadFunction.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New UpFileSave
KSCls.Kesion()
Set KSCls = Nothing

Class UpFileSave
        Private KS,KSUser,FileTitles
		Dim FilePath,MaxFileSize,AllowFileExtStr,AutoReName,RsConfigObj
		Dim FormName,Path,BasicType,ChannelID,UpType,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName
		Dim UpFileObj,FsoObjName,AddWaterFlag,T,CurrNum,CreateThumbsFlag,FieldName,U_FileSize,BoardID,LoginTF
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set T=New Thumb
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set T=Nothing
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Sub Kesion()
		 LoginTF=Cbool(KSUser.UserLoginChecked)
		 IF LoginTF=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		 End If
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"user_upfile.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"selectphoto.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"batchuploadform.asp")<=0 then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 end if
			
        If Cbool(KSUser.UserLoginChecked)=True Then
         IF KS.GetFolderSize(KSUser.GetUserFolder(ksuser.username))/1024>=KS.ChkClng(KSUser.GetUserInfo("SpaceSize")) Then
		  Response.Write "<script>alert('上传失败，您的可用空间不够！');history.back();</script>"
		  response.end
		 End If
		End If
		Response.Write("<style type='text/css'>" & vbcrlf)
		Response.Write("<!--" & vbcrlf)
		Response.Write("body {" & vbcrlf)
		Response.Write("	margin-left: 0px;" & vbcrlf)
		Response.Write("	margin-top: 0px;" & vbcrlf)
		Response.Write("	font-size: 12px;" & vbcrlf)
		'Response.Write("    background:#EEF8FE;" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("-->" & vbcrlf)
		Response.Write("</style>" & vbcrlf)
		
		FsoObjName=KS.Setting(99)
		
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData

		AutoReName = UpFileObj.Form("AutoRename")
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- 图片中心上传 3--下载中心缩略图/文件 41--动漫中心缩略图 42--动漫中心的动漫文件
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("Type")
		BoardID=KS.ChkClng(UpFileObj.Form("BoardID"))
		 
		
		IF BasicType=0 and UpType<>"Field" then 
			Response.Write "<script>alert('请不要非法上传！');history.back();</script>"
			Response.end
		End If
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		AddWaterFlag = UpFileObj.Form("AddWaterFlag")
		If AddWaterFlag <> "1" Then	'生成是否要添加水印标记
			AddWaterFlag = "0"
		End if
		
		'设置文件上传限制,类型及大小
		If UpType="Field" Then
		   Dim RS
		   If ChannelID=0 Then
		   Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_FormField Where FieldID=" & KS.ChkClng(UpFileObj.Form("FieldID")))
		   Else
		   Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_Field Where FieldID=" & KS.ChkClng(UpFileObj.Form("FieldID")))
		   End if
		   If Not RS.Eof Then
		    FieldName=RS(0):MaxFileSize=RS(2):AllowFileExtStr=RS(1)
			FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		   Else
		    Response.End()
		   End IF
		   RS.Close:Set RS=Nothing
		Else
			Select Case BasicType
			  Case 1     '文章中心缩略图
				if Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				If UpType="File" Then '附件
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				Else
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If
			  Case 2     '图片中心上传图片
				 if Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			  Case 3    
				 If Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				If UpType="Pic" Then '下载中心缩略图
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName)& Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				Else
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & "DownUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If
			  Case 4    
				 If Not KS.ReturnChannelAllowUserUpFilesTF(4) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				If UpType="Pic" Then '动漫中心缩略图
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,1)
					FormPath = KS.ReturnChannelUserUpFilesDir(4,KSUser.UserName) & "FlashPhoto/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				Else
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,2)  '取允许上传的动漫类型
					FormPath = KS.ReturnChannelUserUpFilesDir(4,KSUser.UserName) & "FlashUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If
			 Case 5
			     If Not KS.ReturnChannelAllowUserUpFilesTF(5) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				 If UpType="File" Then
				    CreateThumbsFlag=false
				 	MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '设定文件上传最大字节数
				   AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,0)
				 Else
					CreateThumbsFlag=true
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,1)
				End If
				FormPath = KS.ReturnChannelUserUpFilesDir(5,KSUser.UserName) & "Shop/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			 Case 7   
				 If Not KS.ReturnChannelAllowUserUpFilesTF(7) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '设定文件上传最大字节数
				If UpType="Pic" Then '影片缩略图
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,1)
					FormPath = KS.ReturnChannelUserUpFilesDir(7,KSUser.UserName) & "MoviePhoto/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				Else
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,2) &"|" & KS.ReturnChannelAllowUpFilesType(ChannelID,3) & "|"& KS.ReturnChannelAllowUpFilesType(ChannelID,4)  '取允许上传的动漫类型
					FormPath = KS.ReturnChannelUserUpFilesDir(7,KSUser.UserName) & "MovieUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If
	
			Case 8      '供求中心图片
				if Not KS.ReturnChannelAllowUserUpFilesTF(8) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(8)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(8,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(8,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		    Case 9
				if Not KS.ReturnChannelAllowUserUpFilesTF(9) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(9)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(9,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(9,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			Case 9999   '用户头像
				MaxFileSize = 150    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的图片
				FormPath = KS.ReturnChannelUserUpFilesDir(9999,KSUser.UserName)
			Case 9998   '相册封面 
				MaxFileSize = 50    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
				FormPath = KS.ReturnChannelUserUpFilesDir(9998,KSUser.UserName)
			Case 9997 '相片　
				MaxFileSize = 100    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
				FormPath = KS.ReturnChannelUserUpFilesDir(9997,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			Case 9996 '圈子图片　
				MaxFileSize = 50    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9996,KSUser.UserName)
			Case 9995  '音乐
				MaxFileSize = 50000    '设定文件上传最大字节数
				AllowFileExtStr = "mp3"  '取允许上传的动漫类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9995,KSUser.UserName)
			Case 9992 '问答附件	
			 	 If KS.ASetting(42)<>"1" Then
				   KS.Die "<script>alert('对不起，此频道不允许上传附件！');history.back();</script>"
				ElseIf LoginTF=false or (not KS.IsNul(KS.ASetting(46)) and KS.FoundInArr(KS.ASetting(46),KSUser.GroupID,",")=false) Then
			     KS.Die "<script>alert('对不起,您没有在此频道上传的权限!');history.back();</script>"
                 End If

				MaxFileSize =KS.ChkClng(KS.ASetting(44))    '设定文件上传最大字节数
				AllowFileExtStr = KS.ASetting(43)  '取允许上传的类型
				FormPath = KS.ReturnChannelUserUpFilesDir(9997,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		    Case 9994  '论坛
			    If BoardID=0 Then
				  Response.Write "<script>alert('非法参数!');history.back();</script>"
				  Exit Sub
				End If
				KS.LoadClubBoard
				Dim BNode,BSetting
				Set BNode=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & BoardID &"]") 
				If BNode Is Nothing Then KS.Die "非法参数!"
				BSetting=BNode.SelectSingleNode("@settings").text
				BSetting=BSetting & "$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				BSetting=Split(BSetting,"$")
				If KS.ChkClng(BSetting(36))<>1 Then
				  Response.Write "<script>alert('对不起，系统不允许此频道上传文件,请与网站管理员联系!');history.back();</script>"
				  Exit Sub
				End If
				If  LoginTF=true  and (KS.IsNul(BSetting(17)) Or KS.FoundInArr(BSetting(17),KSUser.GroupID,",")) Then
				    If KS.ChkClng(BSetting(39))<>0 Then
					 If Conn.Execute("select count(1) From KS_UploadFiles Where ClassID=" & BoardID & " and datediff(" & DataPart_D & ",AddDate," & SQLNowString & ")<1 and username='" & KSUser.UserName &"'")(0)>=KS.ChkClng(BSetting(39)) Then
					  Response.Write "<script>alert('对不起，本版面每天每人限制只能上传" & KS.ChkClng(BSetting(39))&"个文件!');history.back();</script>"
					  Exit Sub
					 End If
					End If
					MaxFileSize = KS.ChkClng(BSetting(38))    '设定文件上传最大字节数
					AllowFileExtStr = BSetting(37)  '取允许上传的类型
					FormPath =KS.ReturnChannelUserUpFilesDir(9994,KS.Setting(67))
				Else
				  Response.Write "<script>alert('对不起，您没有在本论坛上传附件的权限!');history.back();</script>"
				  Exit Sub
				End If
		    Case 9993  '写日志附件
			    If KS.ChkClng(KS.SSetting(26))=0 Then
				  Response.Write "<script>alert('对不起，系统不允许此频道上传文件,请与网站管理员联系!');history.back();</script>"
				  Exit Sub
			   ElseIf LoginTF=false or (not KS.IsNul(KS.SSetting(30)) and KS.FoundInArr(KS.SSetting(30),KSUser.GroupID,",")=false) Then
			     KS.Die "<script>alert('对不起,您没有在此频道上传的权限!');history.back();</script>"
			   End If
				MaxFileSize = KS.ChkClng(KS.SSetting(28))    '设定文件上传最大字节数
				AllowFileExtStr = KS.SSetting(27)  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9993,KSUser.UserName)
		    Case 999  '上传中心
				MaxFileSize = 100    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png|swf"  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(999,KSUser.UserName)
			Case Else
			  MaxFileSize=0:AllowFileExtStr=""
			  Response.end
			End Select
        End If
		FormPath=Replace(FormPath,".","")
		IF Instr(FormPath,KS.Setting(3))=0 Then FormPath=KS.Setting(3) & FormPath
		FilePath=Server.MapPath(FormPath) & "\"
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
		
        If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		'ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName)
		ReturnValue=CheckUpFile(KSUser,true,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
		
		if ReturnValue <> "" then
		       ReturnValue = Replace(ReturnValue,"'","\'")
			  Response.Write("<script language=""JavaScript"">")
			  Response.Write("alert('" & ReturnValue & "');")
			  if basictype=999 then
			  Response.Write("window.close();")
			  else
			  Response.Write("history.back(-1);")
			 end if
			  Response.Write("</script>")
		else  
            If UpType="Field" Then
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.getElementById('"& FieldName & "').value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>恭喜，上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?ChannelID=" & ChannelID & "&Type=Field&FieldID=" & UpFileObj.Form("FieldID") &"\'>');")
					  Response.Write("</script>")
					  Response.End()
			End If
			TempFileStr=replace(TempFileStr,"'","\'")
			Select Case BasicType
			   Case 1         '文章中心的上传缩略图
				  Response.Write("<script language=""JavaScript"">")
				  if UpType="File" Then   '上传附件
					  Call AddAnnexToDB(ChannelID,KSUser.UserName,TempFileStr,FileTitles,BoardID)
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
					Else
					  if DefaultThumb=0 then
					   Response.Write("parent.document.myform.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  else
						 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
						  Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
						 Else
						  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						 End If
					  end if 
						If Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")(9)=1 Then
						   Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
						end if
						 Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
				   End If 
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?Channelid=" & ChannelID & "&type=" & UpType & "\'>');")
				  Response.Write("</script>")
			   Case 2          '图片中心的上传图片
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=4&Type=" & UpType & "\'>');")
				  Response.Write("</script>")
			  Case 3    '下载中心缩略图
				  Response.Write("<script language=""JavaScript"">")
				  If UPType="Pic" Then
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Else   '下载中心的文件
				  Response.Write("parent.SetDownUrlByUpLoad('" & replace(TempFileStr,"|","") & "'," & U_FileSize & ");")
				  Response.Write("document.write('<br><br><div align=center>文件上传成功！</div>');")
				  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=" & ChannelID & "&Type=" & UPType &"\'>');")
				  Response.Write("</script>")
			  Case 4         '动漫中心的上传缩略图
				  Response.Write("<script language=""JavaScript"">")
				  If UpType="Pic" Then
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Else
				  Response.Write("parent.document.myform.FlashUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br><br><div align=center>文件上传成功！</div>');")
				  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=4&Type=" & UpType & "\'>');")
				  Response.Write("</script>")
			  Case 5         '商城产品
			          Response.Write("<script language=""JavaScript"">")
					  if UpType="File" Then   '上传附件
						  Call AddAnnexToDB(ChannelID,KSUser.UserName,TempFileStr,FileTitles,BoardID)
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
					  ElseIf UpType="ProImage" Then
						  Response.Write("parent.SetPicUrlByUpLoad('" & TempFileStr &  "','" & ThumbPathFileName & "|');")
						  Response.Write("document.write('<br><br><div align=center><font size=2>图片上传成功！</font></div>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?UpType=ProImage&ChannelID=" & ChannelID & "\'>');")
					  Else
						  if DefaultThumb=0 then
						   Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						   Response.Write("parent.document.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  else
						   Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						   Response.Write("parent.document.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=5&Type=" & UpType & "\'>');")
					  Response.Write("</script>")
			  Case 7         '影片中心的上传缩略图
				  Response.Write("<script language=""JavaScript"">")
				  If UpType="Pic" Then
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Else
				  Response.Write("parent.document.myform.MovieUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br><br><div align=center>文件上传成功！</div>');")
				  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=7&Type=" & UpType & "\'>');")
				  Response.Write("</script>")
			  Case 8         '供求中心的上传缩略图
				  Response.Write("<script language=""JavaScript"">")
				  
				  if DefaultThumb=0 then
				   Response.Write("parent.document.myform.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
				  else
					 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
					  Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
					  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
					 Else
					  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
					 End If
				  end if
				  If KS.C_S(ChannelID,34)=0 Then
					       Response.Write("parent.GQContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
				  Else
						   Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
				  End If
				  'Response.Write("parent.GQContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
				  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?Channelid=8\'>');")
				  Response.Write("</script>")
				  Case 9
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.myform.DownUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>试卷上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?Channelid=9\'>');")
					  Response.Write("</script>")		
			  Case 9999        '用户头像
			      session("urel")=""
				  Response.Write("<script language=""JavaScript"">")
				   Response.Write "alert('恭喜，上传成功！');top.location.href='User_EditInfo.asp?action=face&PhotoUrl=" &replace(TempFileStr,"|","") &"';"
				 ' Response.Write("parent.frames['facecut'].location='facecut.asp?photourl=" & replace(TempFileStr,"|","") & "';")
				  'Response.Write("parent.document.myform.UserFace.value='" & replace(TempFileStr,"|","") & "';")
				  'Response.Write("parent.document.myform.showimages.src='" & replace(TempFileStr,"|","") & "';")
				  'response.write "parent.OpenImgCutWindow(0,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;照片上传成功！');")
				 ' Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9999\'>');")
				  Response.Write("</script>")
			  Case 9998        '相册封面
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9998\'>');")
				  Response.Write("</script>")
			  Case 9997        '相片
				  Dim I,TempFileArr
				  TempFileStr=Left(tempfilestr,len(tempfilestr)-1)
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.PhotoUrls.value='" & TempFileStr & "';")
				  TempFileArr=split(TempFileStr,"|")
				  For I=Lbound(TempFileArr) to Ubound(TempFileArr)
				  Response.Write("try{parent.document.myform.view" & I+1 & ".src='" & TempFileArr(i) & "';}catch(e){}")
				  Next
				  Response.Write("</script>")
				  Response.write("<br><br><br><div><font color=red>恭喜您，照片上传成功！请按发布按钮进行保存。</font></div>")
				  Response.Write("<meta http-equiv='refresh' content='2; url=User_upfile.asp?channelid=9997&action=OK'>")
			  Case 9996        '圈子图片
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.showimages.src='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9998\'>');")
				  Response.Write("</script>")
			  Case 9995        '用户头像
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.Url.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;歌曲上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9995\'>');")
				  Response.Write("</script>")
			  Case 9994,9993,9992        '小论坛,博客,问答
			      Response.Write("<script type=""text/JavaScript"">")
				  Dim UpFileArr,UpTitleArr,KK
				  UpFileArr=split(TempFileStr,"|")
				  UpTitleArr=split(FileTitles,"|")
				  For KK=0 To Ubound(UpFileArr)
				   If  Not KS.IsNUL(UpFileArr(kk)) Then
				    Call AddAnnexToDB(ChannelID,KSUser.UserName,UpFileArr(kk),UpTitleArr(kk),BoardID)
				   End If
				  Next
				  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=BatchUploadForm.asp?Channelid=" & ChannelID & "&type=" & UpType & "&boardid="& boardid&"\'>');")
				  Response.Write("</script>")
			  Case 999
				  Response.Write("<script language=""JavaScript"">"&vbcrlf)
				  Response.Write("parent.location.href='selectphoto.asp?channelid=999';"&vbcrlf)
				  Response.Write("</script>"&vbcrlf)
			  Case else
				 if ReturnValue <> "" then
				  Response.Write("<script language=""JavaScript"">"&vbcrlf)
				  Response.Write("alert('" & ReturnValue & "');"&vbcrlf)
				  Response.Write("dialogArguments.location.reload();"&vbcrlf)
				  Response.Write("close();"&vbcrlf)
				  Response.Write("</script>"&vbcrlf)
				 else
				  Response.Write("<script language=""JavaScript"">"&vbcrlf)
				  Response.Write("dialogArguments.location.reload();"&vbcrlf)
				  Response.Write("close();"&vbcrlf)
				  Response.Write("</script>"&vbcrlf)
				 end if
			End Select
         End If
		Set UpFileObj=Nothing
		End Sub
		
End Class
%> 
