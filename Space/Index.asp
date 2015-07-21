<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceApp.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KSCls
Set KSCls = New SpaceIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SpaceIndex
        Private KS, KSR
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Call CloseConn()
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then KS.Die "<script>alert('对不起，本站点关闭空间站点功能!');window.close();</script>"
		    Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			'ks.die QueryStrings
			If QueryStrings<>"" Then 
			  QueryStrings=KS.UrlDecode(QueryStrings)
			  Dim SApp:Set SApp=New SpaceApp
			  SApp.Show(QueryStrings)
			  If SApp.FoundSpace=false Then KS.Die "<script>alert('该用户没有开通空间!');location.href='" & KS.GetDomain &"';</script>"
			  Set SApp=Nothing
			Else
				Dim FileContent
				Set KSR = New Refresh
				FileContent = KSR.LoadTemplate(KS.SSetting(7))
				FCls.RefreshType = "SpaceINDEX" '设置刷新类型，以便取得当前位置导航等
				FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				If Trim(FileContent) = "" Then FileContent = "空间首页模板不存在!"
				FileContent=KSR.KSLabelReplaceAll(FileContent)
		        Set KSR=Nothing
				KS.Echo FileContent 
		   End If 
		End Sub
		
End Class
%>
