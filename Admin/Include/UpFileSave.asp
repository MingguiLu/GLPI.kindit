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
		Dim DefaultThumb    '�趨�ڼ���Ϊ����ͼ
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
			Response.Write "<script>alert('û�е�¼!');history.back();</script>"
			Response.end
		End If
		
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('�Ƿ��ϴ�1��');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"ks.upfileform.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"upfileform.asp")<=0 then
			Response.Write "<script>alert('�Ƿ��ϴ���');history.back();</script>"
			Response.end
		 end if
		 if IsSelfRefer=false Then
			Response.Write "<script>alert('�벻Ҫ�Ƿ��ϴ���');history.back();</script>"
			Response.end
		 End If
		 
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData
		FormPath=Replace(UpFileObj.Form("Path"),".","") 
		IF Instr(FormPath,KS.Setting(3))=0 Then	FormPath=KS.Setting(3) & FormPath
		Call KS.CreateListFolder(FormPath)       '�����ϴ��ļ����Ŀ¼
		FilePath=Server.MapPath(FormPath) & "\"
		If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- ͼƬ�����ϴ� 3--������������ͼ/�ļ� 41--������������ͼ 42--�������ĵĶ����ļ�
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		
		'�����ļ��ϴ�����,���ͼ���С
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
			   Case 0           'Ĭ���ϴ�����
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(0)  '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(0,0)
			   Case 1     '��������
				CreateThumbsFlag=true
				If UpType="Pic" Then  '��������ͼ
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 2     'ͼƬ����
				CreateThumbsFlag=true
				If UpType="Pic" Then  '��������ͼ
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 21     'ͼƬ�����ϴ�ͼƬ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(2)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(2,1)
			  Case 3  
				If UpType="Pic" Then  '����ͼ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else    '���������ļ�	
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 4   
			   If UpType="Pic" Then  '����ͼ
				 CreateThumbsFlag=true
				 MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '�趨�ļ��ϴ�����ֽ���
				 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,1)
			   ElseIf UpType="Flash"  Then'Flash�ļ�
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,2)  'ȡ�����ϴ��Ķ�������
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,0)
			   End If
			 Case 5     '�̳�����
			   If UpType="Pic" Then  '����ͼ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,1)
			   Else
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,0)
			   End If
			 Case 7    'Ӱ����������ͼ
			   If UpType="Pic" Then  '����ͼ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,1)	
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,0)
			   End iF
			 Case 8
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(8)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(8,1)	
			Case 9     '����ϵͳ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(9)   '�趨�ļ��ϴ�����ֽ���
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
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>��ϲ���ϴ��ɹ���</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=Field&FieldID=" & UpFileObj.Form("FieldID") &"\'>');")
					  Response.Write("</script>")
			Else
			    TempFileStr=replace(TempFileStr,"'","\'")
				Select Case BasicType
				   Case 1         '����
					  Response.Write("<script language=""JavaScript"">")
					   if UpType="File" Then   '�ϴ�����
						  Call AddAnnexToDB(ChannelID,KS.C("AdminName"),TempFileStr,FileTitles,0)
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�����ϴ��ɹ���</font>');")
						Else
								 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '����Ƿ��������ͼ
								  Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
								 Else
								  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
								  response.write "parent.OpenImgCutWindow(0,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
								 End If
							  If Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")(9)=1 Then
								 Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
							 End If
							 
							 Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					End If
		
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				   Case 2          'ͼƬ
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
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					 Else
						  Response.Write("parent.SetDownUrlByUpLoad('" & replace(TempFileStr,"|","") & "'," & U_FileSize & ");")
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�ļ��ϴ��ɹ���</font>');")
					 End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 4         '�������ĵ��ϴ�����ͼ
					If UpType="Pic" Then 
					  SuccessDefaultPhoto
					ElseIf UpType="Flash" Then 'Flash�ļ�
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all.FlashUrl.value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('<br><br><div align=center><font size=2>�ļ��ϴ��ɹ���</font></div>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?ChannelID=4&UpType=Flash\'>');")
					  Response.Write("</script>")
					End If
				 Case 5    '�̳���������ͼ
					  Response.Write("<script language=""JavaScript"">")
					  if UpType="File" Then   '�ϴ�����
					      Call AddAnnexToDB(ChannelID,KS.C("AdminName"),TempFileStr,FileTitles,0)
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�����ϴ��ɹ���</font>');")
					  ElseIf UpType="ProImage" Then
						  Response.Write("parent.SetPicUrlByUpLoad('" & TempFileStr &  "','" & ThumbPathFileName & "|');")
						  Response.Write("document.write('<br><br><div align=center><font size=2>ͼƬ�ϴ��ɹ���</font></div>');")
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
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					  End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=5&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 7    'Ӱ����������ͼ
					  If UpType="Pic" Then 
						  SuccessDefaultPhoto
					  Else
					      Response.Write("<script language=""JavaScript"">")
						  Response.Write("parent.SetMovieUrlByUpLoad('" & replace(TempFileStr,"|","") & "');")
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�ļ��ϴ��ɹ���</font>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
					      Response.Write("</script>")
					  End If	  
				  Case 8
					  Response.Write("<script language=""JavaScript"">")
					  if DefaultThumb=0 then
					   Response.Write("parent.document.all.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
					   Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
					  else
						 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '����Ƿ��������ͼ
						  Response.Write("parent.document.all.PhotoUrl.value='" & ThumbPathFileName & "';")
						  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  'ɾ��ԭͼƬ
						 Else
						  Response.Write("parent.document.all.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						  Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
						 End If
					  end if
					  Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?Channelid=8\'>');")
					  Response.Write("</script>")		
				  Case 9
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all.DownUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�Ծ��ϴ��ɹ���</font>');")
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
		
		'�ϴ�Ĭ��ͼ�ɹ�
		Sub SuccessDefaultPhoto()
	      Response.Write("<script language=""JavaScript"">")
		    if DefaultThumb=0 then
				 Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				 Response.write "parent.OpenImgCutWindow(1,'" & KS.Setting(3) & "','" &replace(TempFileStr,"|","") & "');" &vbcrlf
		    else
				 Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
				 Call KS.DeleteFile(replace(TempFileStr,"|",""))  'ɾ��ԭͼƬ
			end if
		   Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
		   Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
		  Response.Write "</script>"
		End Sub
			
End Class
%> 
