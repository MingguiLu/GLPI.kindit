<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.WebFilesCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Frame
KSCls.Kesion()
Set KSCls = Nothing

Class Frame
        Private KS,KSUser
		Private TopDir
		Private Sub Class_Initialize()
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
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>window.close();</script>"
		  Exit Sub
		End If
		TopDir=KSUser.GetUserFolder(ksuser.username)
		 Call KSUser.Head()
		 Call KSUser.InnerLocation("�ҵ��ļ�����")
		 Call KS.CreateListFolder(TopDir)
		  call showframe()
		  call filelist()
		  response.write "<div style=""padding:8px;color:red"">��ܰ���ѣ�Ϊ���˷����ı���ռ䣬�뼰ʱɾ�����õ��ļ���</div>"
		end sub
		
		sub showframe()
        %>
		


<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="1">
                                <tr>
												<td height='25' align='center'>
												
												<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" style="display:nowrap">
						<tr class="tdbg">
						<td width="157" align="right">�����ܿռ� <font color=red><%=round(KSUser.GetUserInfo("SpaceSize")/1024,2)%>M</font>,ʹ�������</td>
						<td><img src="images/bar.gif" width="0" height="16" id="Sms_bar" align="absmiddle" /></td>
						<td width="211"  align="center" id="Sms_txt">100%</td>
						</tr></table>
		 <%
        response.write showtable("Sms_bar","Sms_txt",KS.GetFolderSize(TopDir)/1024,KSUser.GetUserInfo("SpaceSize"))
%>
												
												</td></tr>
												
                        </table>
		<%
		end sub
		
		sub filelist()
		 Response.Buffer = True
		Response.Expires = -1
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
		Dim WFCls:Set WFCls = New WebFilesCls
		Call WFCls.Kesion(0,TopDir,"",20,"","Images/Css.css")
		Set WFCls = Nothing
	  
      End Sub
	   '��ͼƬ�������ƣ�����������ƣ���������������
		Function ShowTable(SrcName,TxtName,str,c)
		Dim Tempstr,Src_js,Txt_js,TempPercent
		If C = 0 Then C = 99999999
		Tempstr = str/C
		TempPercent = FormatPercent(tempstr,0,-1)
		Src_js = "document.getElementById(""" + SrcName + """)"
		Txt_js = "document.getElementById(""" + TxtName + """)"
			ShowTable = VbCrLf + "<script>"
			ShowTable = ShowTable + Src_js + ".width=""" & FormatNumber(tempstr*600,0,-1) & """;"
			ShowTable = ShowTable + Src_js + ".title=""��������Ϊ��"&c/1024&" MB�����ã�"&FormatNumber(str/1024,2)&"��MB��"";"
			ShowTable = ShowTable + Txt_js + ".innerHTML="""
			If FormatNumber(tempstr*100,0,-1) < 80 Then
				ShowTable = ShowTable + "��ʹ��:" & TempPercent & """;"
			Else
				ShowTable = ShowTable + "<font color=\""red\"">��ʹ��:" & TempPercent & ",��Ͽ�����</font>"";"
			End If
			ShowTable = ShowTable + "</script>"
		End Function
		
End Class
%> 
