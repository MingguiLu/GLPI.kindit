<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/UploadFunction.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
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
        Private KS,KSUser,FileTitles,Title
		Dim FilePath,MaxFileSize,AllowFileExtStr,BasicType,ChannelID,UpType
		Dim FormName,Path,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName
		Dim UpFileObj,CurrNum,CreateThumbsFlag,FieldName,U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
		Sub Kesion()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  KS.Die "error:" & escape("没有登录!")
		End If

		Set UpFileObj = New UpFileClass
		on error resume next
		UpFileObj.GetData
		If ERR.Number<>0 Then Set UpFileObj=Nothing : err.clear:KS.Die "error:" & escape("上传失败，可能您的上传的文件太大!")
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))   
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		Select Case BasicType
		  Case 7     '影片
		    MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
		    AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,2) &"|" & KS.ReturnChannelAllowUpFilesType(ChannelID,3) & "|"& KS.ReturnChannelAllowUpFilesType(ChannelID,4)  '取允许上传的动漫类型
			FormPath = KS.ReturnChannelUserUpFilesDir(7,KSUser.UserName) & "MovieUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			
		  Case 2     '图片中心
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		 Case 9997  '相片
		 	MaxFileSize = 200    '设定文件上传最大字节数
			AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
			FormPath = KS.ReturnChannelUserUpFilesDir(9997,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		End Select
		FormPath=Replace(FormPath,".","")
		IF Instr(FormPath,KS.Setting(3))=0 Then FormPath=KS.Setting(3) & FormPath
		FilePath=Server.MapPath(FormPath) & "\"
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
		
        If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		ReturnValue = CheckUpFile(KSUser,true,UpFileObj,FormPath,FilePath,MaxFileSize,AllowFileExtStr,U_FileSize,TempFileStr,FileTitles,CurrNum,CreateThumbsFlag,DefaultThumb,ThumbPathFileName)
		if ReturnValue <> "" then
		     KS.Die "error:" & escape(Replace(replace(ReturnValue,"\n",""),"'","\'"))
		else 
			 TempFileStr=replace(TempFileStr,"'","\'")
			 Select Case BasicType
			      Case 7
				      KS.Die replace(TempFileStr,"|","")
				  Case 2         '图片
					   KS.Die replace(TempFileStr,"|","") &  "@" & ThumbPathFileName & "@"
				  Case 9997    '相片
					   KS.Die replace(TempFileStr,"|","") &  "@" & replace(TempFileStr,"|","") & "@"
			 End Select
		  End iF
		Set UpFileObj=Nothing
 End Sub
End Class

%> 
