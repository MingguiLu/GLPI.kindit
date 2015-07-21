<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312" 
Dim KSCls
Set KSCls = New UserAjax
KSCls.Kesion()
Set KSCls = Nothing

Class UserAjax
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		  Select Case KS.S("Action")
		   Case "GetNewMessage" Call GetNewMessage()
		   Case "GetAdminMessage" Call GetAdminMessage()
		   Case "TalkSave" Call TalkSave()
		  End Select
		End Sub
		
		Sub GetNewMessage()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write Escape("站内消息(0)")
		  Exit Sub
		End If
		Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
		'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogMessage Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
		'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogComment Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
		'MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_Friend Where Friend='" &KSUser.UserName &"' And accepted=0")(0)

		Response.write Escape("站内消息(<font color='#ff0000'>" & MyMailTotal&"</font>)")
		If MyMailTotal>0 Then Response.Write "<bgsound src=""images/mail.wmv"" border=0>"
		End Sub

		Sub GetAdminMessage()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "(0)"
		  Exit Sub
		End If
		Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
		Response.write "(<font color='#ff0000'>" & MyMailTotal&"</font>)"
		If MyMailTotal>0 Then Response.Write "<bgsound src=""../User/images/mail.wmv"" border=0>"
		End Sub
		
		Sub TalkSave()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		   KS.Die Escape("请先登录！")
		 End If
		 Dim Content:Content=UnEscape(KS.S("Content"))
		 If Len(Content)<10 Then KS.Die Escape("多说几个字吧！")
		 If KS.IsNul(Content) Then KS.Die Escape("请输入内容！")
		 Dim Title:Title=Left(Content,200)
		 Dim TypeID:TypeID=KS.ChkClng(Conn.Execute("Select Top 1 TypeID From KS_BlogType Where IsDefault=1")(0))
	 	 Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
		  RSObj.Open "Select top 1 * From KS_BlogInfo",Conn,1,3
		   RSObj.AddNew
		    RSObj("IsTalk")=1
			RSObj("Hits")=0
			RSObj("UserID")=KSUser.GetUserInfo("userid")
			RSObj("Title")=Title
			RSObj("Content")=Content
			RSObj("TypeID")=TypeID
			if KS.ChkClng(KS.SSetting(3))=1 Then
			  RSObj("Status")=2
			Else
			  RSObj("Status")=0
			End if
			 RSObj("UserName")=KSUser.UserName
			 RSObj("AddDate")=Now
			 RSObj("face")=1
			 RSObj("Weather")="sun.gif"
		   RSObj.Update
		  RSObj.Close :  Set RSObj=Nothing
		  Call KSUser.AddLog(KSUser.UserName,Content & " <a href=""{$GetSiteUrl}space/morefresh.asp"" target=""_blank"">[新鲜事]</a>",100)
		  KS.Die "success"
	End Sub
End Class
%> 
