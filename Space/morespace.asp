<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 7.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KSCls
Set KSCls = New Spacemore
KSCls.Kesion()
Set KSCls = Nothing

Class Spacemore
        Private KS, KSR,CurrPage,MaxPerPage,TotalPut,str
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		      Dim FileContent
				   FileContent = KSR.LoadTemplate(KS.SSetting(8))
				   FCls.RefreshType = "MoreSpace" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   Application(KS.SiteSN & "RefreshFolderID") = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   If Trim(FileContent) = "" Then FileContent = "�ռ丱ģ�岻����!"
				   FileContent=Replace(FileContent,"{$ShowMain}",GetSpaceList())
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
		           Response.Write FileContent  
		End Sub
	%>
	<!--#Include file="../ks_cls/ubbfunction.asp"-->
	<%	
		
 '�ռ��б�
 Function GetSpaceList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(9))
		 dim classid:classid=ks.chkclng(ks.s("classid"))
		 dim recommend:recommend=ks.chkclng(ks.s("recommend"))
		 CurrPage = KS.ChkClng(KS.G("page"))
		 If CurrPage<=0 Then CurrPage = 1
		 
	    dim rsc:set rsc=conn.execute("select classname,classid from ks_blogclass order by orderid")
	   if not rsc.eof then
	   str="<div class=""categorybox"">" & vbcrlf
	   str=str &"<ul><li>����鿴��</li>"
		   If classid=0 then 
		     str=str &"<li class=""curr""><a href='morespace.asp'>���з���</a></li>"
		   else
		     str=str &"<li><a href='morespace.asp'>���з���</a></li>"
		   end if
	    do while not rsc.eof
		 if classid=rsc(1) then
		   str=str & "<li class=""curr""><a href='?classid=" & rsc(1) &"'>" & rsc(0) & "</a></li>"
		 else
		   str=str & "<li><a href='?classid=" & rsc(1) &"'>" & rsc(0) & "</a></li>"
		 end if
		 rsc.movenext
		loop
	   end if
	   rsc.close:set rsc=nothing
   str=str &"</ul>" & vbcrlf
   str=str &"</div>" &vbcrlf	 
		 
  str=str & "<table border=""0"" cellpadding=""1"" cellspacing=""1"" width=""98%"" backcolor=""#efefef"">"
 
 dim param:param=" where status=1"
 if classid<>0 then param=param & " and a.classid=" & classid
 if recommend<>0 then param=param & " and recommend=1"

 if ks.s("key")<>"" then param=param & " and blogname like '%" & ks.r(ks.s("key")) &"%'"
    Dim SQLStr:SQLStr="select a.*,b.classname,u.userface,u.realname from (ks_blog a inner join ks_blogclass b on a.classid=b.classid) inner join ks_user u on a.username=u.username " & param & " order by a.hits desc,a.blogid desc"
	Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		rsobj.open SQLStr ,conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 	str=str & "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=4><p>�Բ���û���ҵ��ռ�! </p></td></tr>"
				 Else
							  totalPut = conn.execute("select count(1) from (ks_blog a inner join ks_blogclass b on a.classid=b.classid) inner join ks_user u on a.username=u.username " & param)(0)
								If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
								Else
										CurrPage = 1
								End If
								call ShowSpaceList(RSObj)
				           End If
		 
		 str=str &  "            </table>" & vbcrlf
		 str=str & KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
		 RSObj.Close:Set RSObj=Nothing
		 
		 str=str & "<div class=""clear""></div><table>"
		  str=str & "<form name=""myform"" action=""morespace.asp"" method=""get""/> <tr height=""22"">"
	   str=str & "<td align=""center"" colspan=2><strong>���ռ�����������</strong><input style=""border:1px #000 solid;height:18px;"" type=""text"" size=""12"" name=""key"">&nbsp;&nbsp;<input type=""submit"" value= "" �� �� "" class=""btn""></td>"
	   str=str & "</form></tr>"
	   str=str & "</table><br/><br/>"
		 GetSpaceList=str
  End Function

  Sub ShowSpaceList(rs)
   dim i,logo,rss
   do while not rs.eof
     logo=RS("Logo")
	 If KS.IsNul(Logo) Then
	   logo=RS("UserFace")
	 End If
	 If KS.IsNul(Logo) Then Logo="images/face/boy.jpg"
	 If Left(logo,1)<>"/" and Left(lcase(logo),4)<>"http" Then Logo=KS.Setting(3) & Logo
	 str=str & "<tr>"
      str=str & "<td class=""mysplittd"" style=""width:40px""><img title=""����ʱ�䣺" & rs("adddate") & """ style=""border:1px solid #efefef;padding:2px"" src=""" & Logo & """ width=""52"" height=""52"" /></td><td class=""mysplittd"">"
		  dim spacedomain,predomain
		  If KS.SSetting(14)="1"  Then
		   predomain=rs("domain")
		  end if
		  if predomain<>"" then
		   spacedomain="http://" & predomain & "." & KS.SSetting(16)
		  else
		    spacedomain=KS.GetSpaceUrl(rs("userid"))
		  end if

      str=str & "<a title=""" & rs("blogname") & """ href=""" & spacedomain  &""" target=""blank""> " & rs("blogname")  &"</a>"
	  if rs("recommend")=1 then str=str & "<font color=red>[��]</font>"
	  str=str &" ���ࣺ" & rs("classname")
	  str=str & "<div class=""intro""> " & rs("Descript") & "</div>"
	  
	  set rss=conn.execute("select top 1 * From KS_BlogInfo Where UserName='" & RS("UserName") & "' and status=0")
	  If Not RSS.Eof Then
	   str=str &"<div class=""fresh"">" & KSR.ReplaceEmot(Ubbcode(rss("content"),0)) 
	   If RSS("Istalk")="1" Then
	    str=str& "<a href='" & spacedomain & "/log/" & rss("id") & "' target='_blank'>[����]</a>"
	   Else
	    str=str& "<a href='" & spacedomain & "/fresh' target='_blank'>[������]</a>"
	   End If
	    str=str &"- <a href='" & spacedomain & "/log/" & rss("id") & "' target='_blank'>����(<font color=red>" & rss("totalput") & "</font>)</a></div>"
	  End If
	  RSS.Close
	  str=str & "<div class=""btntips""><a href='javascript:void(0)' onclick=""addF(event,'" & rs("username") & "')"">��Ϊ����</a> | <a href='javascript:void(0)' onclick=""sendMsg(event,'"& rs("username") & "')"">������Ϣ</a> | <a href='" & SpaceDomain & "' target='_blank'>��ע��" & rs("UserName") & "�� �Ŀռ�</a></div>"
	  str=str & "</td>"
	  str=str & "</tr>"
   rs.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
	  loop
  End Sub		
		
		
End Class
%>
