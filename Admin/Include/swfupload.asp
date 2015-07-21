<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../../KS_Cls/UploadFunction.asp"-->
<!--#include file="Session.asp"-->
<%
Server.ScriptTimeout=9999999
Response.CharSet="gb2312"
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
        Private KS,FileTitles,Title
		Dim FilePath,MaxFileSize,AllowFileExtStr,BasicType,ChannelID,UpType
		Dim FormName,Path,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName
		Dim UpFileObj,CurrNum,CreateThumbsFlag,FieldName,	U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		Sub Kesion()
		 If KS.IsNul(KS.C("AdminName")) Or KS.IsNul(KS.C("AdminPass")) Or KS.IsNul(KS.C("PowerList"))="" Or KS.IsNUL(KS.C("UserName")) Then
			KS.Die "error:" & escape("没有登录!")
		End If

		Set UpFileObj = New UpFileClass
		on error resume next
		UpFileObj.GetData
		If ERR.Number<>0 Then err.clear:KS.Die "error:" & escape("上传失败，可能您的上传的文件太大!")
		FormPath=KS.GetUpFilesDir
		FilePath=Server.MapPath(FormPath) & "\"
		FormPath=FormPath & "/"
		If KS.Setting(97)=1 Then FormPath=KS.Setting(2) & FormPath
		
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))   
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		
		Select Case BasicType
		  Case 2,5     '图片中心
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
		End Select
			
		ReturnValue = CheckUpFile("",false,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
		if ReturnValue <> "" then
		     ReturnValue=replace(ReturnValue,"\n","。")
		     If Instr(ReturnValue,"上传失败")<>0 Then
		     KS.Die "error:" & escape("上传失败" & Replace(Split(ReturnValue,"上传失败")(1),"'","\'"))
			 Else
		     KS.Die "error:" & escape(Replace(ReturnValue,"'","\'"))
			 End If
		else 
			 TempFileStr=replace(TempFileStr,"'","\'")
			 Select Case BasicType
				  Case 2,5          '图片
					  KS.Die replace(TempFileStr,"|","") &  "@" & ThumbPathFileName & "@"
					  'KS.Die replace(TempFileStr,"|","") &  "@" & ThumbPathFileName & "@" & escape(FileTitles)
			 End Select
		  End iF
		Set UpFileObj=Nothing
 End Sub
End Class

%> 
