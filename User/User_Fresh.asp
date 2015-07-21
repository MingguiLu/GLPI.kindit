<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KS,KSUser
Set KS=New PublicCls
Set KSUser = New UserCls

If KS.S("Action")="delfresh" Then DelFresh
Set KSUser=Nothing
Set KS=Nothing
CloseConn     

'删除新鲜事
Sub DelFresh()
	  Dim KSUser:Set KSUser=New UserCls
	  If Cbool(KSUser.UserLoginChecked)=false Then
	    KS.AlertHintScript "对不起，没有权限操作!"
		Exit Sub
	  End If
	  Dim ID:ID=KS.ChkClng(KS.S("ID"))
	  If ID=0 Then 
	    KS.AlertHintScript "出错啦!"
		Exit Sub
	  End If
	  Dim RS:Set RS=Conn.Execute("select id From KS_BlogInfo Where ID=" & id & " And UserName='" & KSUser.UserName & "'")
	  Do While Not RS.Eof
	  Conn.Execute("Delete From KS_BlogComment Where LogID=" & RS("id"))
	  RS.MoveNext
	  Loop
	  RS.Close
	  Set RS=Nothing
	  Conn.Execute("Delete From KS_BlogInfo Where ID=" & id & " And UserName='" & KSUser.UserName & "'")
	  KS.AlertHintScript "恭喜，删除成功！"
 End Sub

%> 