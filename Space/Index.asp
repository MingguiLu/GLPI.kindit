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
		    If KS.SSetting(0)=0 Then KS.Die "<script>alert('�Բ��𣬱�վ��رտռ�վ�㹦��!');window.close();</script>"
		    Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			'ks.die QueryStrings
			If QueryStrings<>"" Then 
			  QueryStrings=KS.UrlDecode(QueryStrings)
			  Dim SApp:Set SApp=New SpaceApp
			  SApp.Show(QueryStrings)
			  If SApp.FoundSpace=false Then KS.Die "<script>alert('���û�û�п�ͨ�ռ�!');location.href='" & KS.GetDomain &"';</script>"
			  Set SApp=Nothing
			Else
				Dim FileContent
				Set KSR = New Refresh
				FileContent = KSR.LoadTemplate(KS.SSetting(7))
				FCls.RefreshType = "SpaceINDEX" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				If Trim(FileContent) = "" Then FileContent = "�ռ���ҳģ�岻����!"
				FileContent=KSR.KSLabelReplaceAll(FileContent)
		        Set KSR=Nothing
				KS.Echo FileContent 
		   End If 
		End Sub
		
End Class
%>
