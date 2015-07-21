<% Option Explicit %>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/UploadFunction.asp"-->
<!--#include file="Session.asp"-->
<%
Server.ScriptTimeout=9999999
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
        Private KS,FileTitles
		Dim FilePath,MaxFileSize,AllowFileExtStr,BasicType,ChannelID,UpType
		Dim FormName,Path,TempFileStr,FormPath,ThumbPathFileName
		Dim UpFileObj,CurrNum,CreateThumbsFlag,FieldName,U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		Function IsSelfRefer() 
			Dim sHttp_Referer, sServer_Name 
			sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER")) 
			sServer_Name = CStr(Request.ServerVariables("SERVER_NAME")) 
			If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then 
			IsSelfRefer = True 
			Else 
			IsSelfRefer = False 
			End If 
		End Function 
		Sub Kesion()
		Response.Write("<style type='text/css'>" & vbcrlf)
		Response.Write("<!--" & vbcrlf)
		Response.Write("body {background:#f0f0f0;" & vbcrlf)
		Response.Write("	margin-left: 0px;" & vbcrlf)
		Response.Write("	margin-top: 0px;" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("-->" & vbcrlf)
		Response.Write("</style>" & vbcrlf)
		
		If KS.IsNul(KS.C("AdminName")) Or KS.IsNul(KS.C("AdminPass")) Or KS.IsNul(KS.C("PowerList"))="" Or KS.IsNUL(KS.C("UserName")) Then
			Response.Write "<script>alert('没有登录!');history.back();</script>"
			Response.end
		End If
		
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('非法上传1！');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"ks.upfileform.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"upfileform.asp")<=0 then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 end if
		 if IsSelfRefer=false Then
			Response.Write "<script>alert('请不要非法上传！');history.back();</script>"
			Response.end
		 End If
		 
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData
		FormPath=Replace(UpFileObj.Form("Path"),".","") 
		IF Instr(FormPath,KS.Setting(3))=0 Then	FormPath=KS.Setting(3) & FormPath
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
		FilePath=Server.MapPath(FormPath) & "\"
		If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- 图片中心上传 3--下载中心缩略图/文件 41--动漫中心缩略图 42--动漫中心的动漫文件
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		
		'设置文件上传限制,类型及大小
		If UpType="Field" Then
		   Dim RS:Set RS=Conn.Execute("Select top 1 FieldName,AllowFileExt,MaxFileSize From KS_Field Where FieldID=" & KS.ChkClng(UpFileObj.Form("FieldID")))
		   If Not RS.Eof Then
		    FieldName=RS(0):MaxFileSize=RS(2):AllowFileExtStr=RS(1)
		   Else
		    Response.End()
		   End IF
		   RS.Close:Set RS=Nothing
		Else
			Select Case BasicType
			   Case 0           '默认上传参数
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(0)  '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(0,0)
			   Case 1     '文章中心
				CreateThumbsFlag=true
				If UpType="Pic" Then  '文章缩略图
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 2     '图片中心
				CreateThumbsFlag=true
				If UpType="Pic" Then  '文章缩略图
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 21     '图片中心上传图片
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(2)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(2,1)
			  Case 3  
				If UpType="Pic" Then  '缩略图
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else    '下载中心文件	
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 4   
			   If UpType="Pic" Then  '缩略图
				 CreateThumbsFlag=true
				 MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,1)
			   ElseIf UpType="Flash"  Then'Flash文件
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,2)  '取允许上传的动漫类型
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,0)
			   End If
			 Case 5     '商城中心
			   If UpType="Pic" Then  '缩略图
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,1)
			   Else
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,0)
			   End If
			 Case 7    '影视中心缩略图
			   If UpType="Pic" Then  '缩略图
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,1)	
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,0)
			   End iF
			 Case 8
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(8)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(8,1)	
			Case 9     '考试系统
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(9)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(9,0)
			End Select
		End If
			
		ReturnValue = CheckUpFile("",false,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
		if ReturnValue <> "" then
		     ReturnValue = Replace(ReturnValue,"'","\'")
		     KS.AlertHintScript ReturnValue
			 Response.End()
		else 
			If UpType="Field" Then
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all."& FieldName & ".value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>恭喜，上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=Field&FieldID=" & UpFileObj.Form("FieldID") &"\'>');")
					  Response.Write("</script>")
			Else
			    TempFileStr=replace(TempFileStr,"'","\'")
				Select Case BasicType
				   Case 1         '文章
					  Response.Write("<script language=""JavaScript"">")
					   if UpType="File" Then   '上传附件
						  Call AddAnnexToDB(ChannelID,KS.C("AdminName"),TempFileStr,FileTitles,0)
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
						Else
								 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
								  Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
								 Else
								  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
								  response.write "parent.OpenImgCutWindow(0,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
								 End If
							  If Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")(9)=1 Then
								 Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
							 End If
							 
							 Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					End If
		
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				   Case 2          '图片
					  SuccessDefaultPhoto
				  Case 3  
					  Response.Write("<script language=""JavaScript"">")
					 If UpType="Pic" Then
					  if DefaultThumb=0 then
					   Response.Write("parent.document.getElementById('PhotoUrl').value='" & replace(TempFileStr,"|","") & "';")
					   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
					   response.write "parent.OpenImgCutWindow(0,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
					  else
					   Response.Write("parent.document.getElementById('PhotoUrl').value='" & ThumbPathFileName & "';")
					   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
					  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					 Else
						  Response.Write("parent.SetDownUrlByUpLoad('" & replace(TempFileStr,"|","") & "'," & U_FileSize & ");")
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>文件上传成功！</font>');")
					 End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 4         '动漫中心的上传缩略图
					If UpType="Pic" Then 
					  SuccessDefaultPhoto
					ElseIf UpType="Flash" Then 'Flash文件
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all.FlashUrl.value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('<br><br><div align=center><font size=2>文件上传成功！</font></div>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?ChannelID=4&UpType=Flash\'>');")
					  Response.Write("</script>")
					End If
				 Case 5    '商城中心缩略图
					  Response.Write("<script language=""JavaScript"">")
					  if UpType="File" Then   '上传附件
					      Call AddAnnexToDB(ChannelID,KS.C("AdminName"),TempFileStr,FileTitles,0)
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
					  ElseIf UpType="ProImage" Then
						  Response.Write("parent.SetPicUrlByUpLoad('" & TempFileStr &  "','" & ThumbPathFileName & "|');")
						  Response.Write("document.write('<br><br><div align=center><font size=2>图片上传成功！</font></div>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?UpType=ProImage&ChannelID=" & ChannelID & "\'>');")
					  Else
						  if DefaultThumb=0 then
						   Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						   Response.Write("parent.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
					       Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
						  else
						   Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						   Response.Write("parent.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=5&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 7    '影视中心缩略图
					  If UpType="Pic" Then 
						  SuccessDefaultPhoto
					  Else
					      Response.Write("<script language=""JavaScript"">")
						  Response.Write("parent.SetMovieUrlByUpLoad('" & replace(TempFileStr,"|","") & "');")
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>文件上传成功！</font>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
					      Response.Write("</script>")
					  End If	  
				  Case 8
					  Response.Write("<script language=""JavaScript"">")
					  if DefaultThumb=0 then
					   Response.Write("parent.document.all.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
					   Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
					  else
						 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
						  Response.Write("parent.document.all.PhotoUrl.value='" & ThumbPathFileName & "';")
						  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
						 Else
						  Response.Write("parent.document.all.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						  Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
						 End If
					  end if
					  Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?Channelid=8\'>');")
					  Response.Write("</script>")		
				  Case 9
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all.DownUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>试卷上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?Channelid=9\'>');")
					  Response.Write("</script>")		
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
		  End iF
		Set UpFileObj=Nothing
		End Sub
		
		'上传默认图成功
		Sub SuccessDefaultPhoto()
	      Response.Write("<script language=""JavaScript"">")
		    if DefaultThumb=0 then
				 Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				 Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
		    else
				 Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
				 Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
			end if
		   Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
		   Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
		  Response.Write "</script>"
		End Sub
			
End Class
%> 
